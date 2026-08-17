import time


class CircuitBreaker:
    """Tracks consecutive flush failures and pauses API calls during an outage.

    In-process only (resets on restart) — that's acceptable here since the
    worker thread already re-derives the current interval/backoff on boot,
    and a restart is itself a reasonable recovery action for a stuck sync.
    """

    def __init__(self, failure_threshold=5, cooldown_seconds=300):
        self.failure_threshold = failure_threshold
        self.cooldown_seconds = cooldown_seconds
        self._consecutive_failures = 0
        self._opened_at = None

    def record_success(self):
        self._consecutive_failures = 0
        self._opened_at = None

    def record_failure(self):
        self._consecutive_failures += 1
        if self._consecutive_failures >= self.failure_threshold and self._opened_at is None:
            self._opened_at = time.monotonic()

    @property
    def is_open(self):
        if self._opened_at is None:
            return False
        if time.monotonic() - self._opened_at >= self.cooldown_seconds:
            # Cooldown elapsed: allow one probe attempt through.
            self._opened_at = None
            self._consecutive_failures = self.failure_threshold - 1
            return False
        return True

    @property
    def consecutive_failures(self):
        return self._consecutive_failures
