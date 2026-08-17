import logging
import os
import queue
import sys
import threading
import time

from django.conf import settings
from django.db import connection

logger = logging.getLogger(__name__)

_change_queue = queue.Queue(maxsize=20000)
_worker_running = False
_worker_instance = None


def get_queue():
    return _change_queue


def get_worker_status():
    """Read-only snapshot for the status dashboard. Same-process only —
    each Django process (e.g. a management command) has its own worker, so
    this only reflects the process serving the request."""
    if _worker_instance is None or _worker_instance._breaker is None:
        return {'running': False}
    breaker = _worker_instance._breaker
    return {
        'running': True,
        'queue_size': _change_queue.qsize(),
        'circuit_open': breaker.is_open,
        'consecutive_failures': breaker.consecutive_failures,
    }


def _col_letter(n):
    letters = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


class SheetsSyncWorkerThread(threading.Thread):
    def __init__(self):
        super().__init__()
        self.daemon = True
        self.name = "SheetsSyncWorker"
        self._breaker = None

    def run(self):
        global _worker_running
        _worker_running = True
        logger.info("Sheets sync worker thread started.")
        time.sleep(10)

        from sheets_sync.circuit_breaker import CircuitBreaker

        settings_obj = self._get_settings()
        self._breaker = CircuitBreaker(
            failure_threshold=(settings_obj.circuit_breaker_failure_threshold if settings_obj else 5),
            cooldown_seconds=(settings_obj.circuit_breaker_cooldown_seconds if settings_obj else 300),
        )
        interval = (settings_obj.flush_interval_seconds if settings_obj else 4) or 4

        while _worker_running:
            time.sleep(interval)
            try:
                connection.close_if_unusable_or_obsolete()
                settings_obj = self._get_settings()
                base_interval = (settings_obj.flush_interval_seconds if settings_obj else 4) or 4

                if not settings_obj or not settings_obj.enabled:
                    interval = base_interval
                    continue

                if self._breaker.is_open:
                    interval = base_interval
                    continue

                events = self._drain(settings_obj.max_batch_size)
                if not events:
                    interval = base_interval
                    continue

                all_ok = self._flush(events)
                if all_ok:
                    self._breaker.record_success()
                    interval = base_interval
                else:
                    self._breaker.record_failure()
                    interval = min(base_interval * (2 ** self._breaker.consecutive_failures), 300)
            except Exception:
                logger.exception("Error in sheets sync worker tick")
                if self._breaker:
                    self._breaker.record_failure()
                interval = min(interval * 2, 300) if interval else 8

    def _get_settings(self):
        from sheets_sync.models import SheetsSyncSetting
        try:
            return SheetsSyncSetting.get_settings()
        except Exception:
            return None

    def _drain(self, max_batch_size):
        events = []
        for _ in range(max_batch_size):
            try:
                events.append(_change_queue.get_nowait())
            except queue.Empty:
                break

        # Dedupe: keep only the latest event per (tab, pk) -- last-write-wins,
        # which is correct for a DR mirror since only final state matters.
        deduped = {}
        for event in events:
            deduped[(event.tab_name, event.object_pk)] = event
        return list(deduped.values())

    def _flush(self, events):
        from sheets_sync import client as sheets_client
        from sheets_sync.models import SheetsRowIndex, SheetsSyncLog

        by_tab = {}
        for event in events:
            by_tab.setdefault(event.tab_name, []).append(event)

        try:
            spreadsheet = sheets_client.open_spreadsheet()
        except Exception as exc:
            logger.error("sheets_sync: could not open spreadsheet: %s", exc)
            for tab_name, tab_events in by_tab.items():
                SheetsSyncLog.objects.create(
                    tab_name=tab_name, batch_size=len(tab_events), status='FAILED',
                    error_message=str(exc)[:2000],
                )
                self._requeue(tab_events)
            return False

        all_ok = True
        for tab_name, tab_events in by_tab.items():
            start = time.monotonic()
            try:
                headers = tab_events[0].headers
                worksheet = sheets_client.get_or_create_worksheet(spreadsheet, tab_name, headers)

                pks = [e.object_pk for e in tab_events]
                existing = {
                    row.object_pk: row.row_number
                    for row in SheetsRowIndex.objects.filter(tab_name=tab_name, object_pk__in=pks)
                }

                last_col_letter = _col_letter(len(headers))
                updates = []
                new_index_rows = []

                last_known = SheetsRowIndex.objects.filter(tab_name=tab_name).order_by('-row_number').first()
                next_free_row = (last_known.row_number + 1) if last_known else 2

                for event in tab_events:
                    row_number = existing.get(event.object_pk)
                    if row_number is None:
                        row_number = next_free_row
                        next_free_row += 1
                        new_index_rows.append(SheetsRowIndex(
                            tab_name=tab_name, object_pk=event.object_pk, row_number=row_number,
                        ))
                    updates.append({
                        'range': f"A{row_number}:{last_col_letter}{row_number}",
                        'values': [event.row],
                    })

                if worksheet.row_count < next_free_row:
                    worksheet.add_rows(next_free_row - worksheet.row_count + 50)

                # Persist the row index BEFORE the API call: if the write then
                # fails, the retry will resolve these pks as "existing" and
                # issue an idempotent update instead of appending duplicates.
                if new_index_rows:
                    SheetsRowIndex.objects.bulk_create(new_index_rows, ignore_conflicts=True)

                worksheet.batch_update(updates, value_input_option='RAW')

                SheetsSyncLog.objects.create(
                    tab_name=tab_name, batch_size=len(tab_events), status='SUCCESS',
                    duration_ms=int((time.monotonic() - start) * 1000),
                )
            except Exception as exc:
                logger.exception("sheets_sync: flush failed for tab %s", tab_name)
                SheetsSyncLog.objects.create(
                    tab_name=tab_name, batch_size=len(tab_events), status='FAILED',
                    error_message=str(exc)[:2000],
                    duration_ms=int((time.monotonic() - start) * 1000),
                )
                self._requeue(tab_events)
                all_ok = False

        return all_ok

    def _requeue(self, events):
        for event in events:
            try:
                _change_queue.put_nowait(event)
            except Exception:
                logger.warning("sheets_sync: could not requeue event for %s after failure", event.tab_name)


def start_worker():
    global _worker_instance
    # Never run under the test runner.
    if 'test' in sys.argv or any('test' in arg for arg in sys.argv):
        logger.info("Running under test runner. Skipping sheets sync worker start.")
        return

    # Opt-out switch, defaults ON so it works the same way in dev/prod.
    if getattr(settings, 'SHEETS_SYNC_INPROCESS_WORKER', True) is False:
        logger.info("In-process sheets sync worker disabled via settings. Skipping.")
        return

    # Avoid a duplicate worker under Django's autoreloader (watcher + child).
    running_under_reloader = ('runserver' in sys.argv) and ('--noreload' not in sys.argv)
    if running_under_reloader and os.environ.get('RUN_MAIN') != 'true':
        logger.info("Reloader parent detected; the child process will start the sheets sync worker.")
        return

    for thread in threading.enumerate():
        if thread.name == "SheetsSyncWorker":
            logger.info("Sheets sync worker thread is already running.")
            return

    _worker_instance = SheetsSyncWorkerThread()
    _worker_instance.start()
    logger.info("In-process sheets sync worker started.")
