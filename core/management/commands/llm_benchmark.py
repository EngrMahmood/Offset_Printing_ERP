"""Benchmark candidate LLM models against real facts blocks used by this app.

Measures wall-clock time, token usage, and a faithfulness check (do all the
expected facts survive, verbatim, into the model's answer) — so choosing a
model is based on this app's actual prompts, not generic benchmarks.
"""
import time

from django.core.management.base import BaseCommand

from core.llm.client import call_chat
from core.llm.prompts import CHAT_ASSISTANT_SYSTEM_PROMPT, NARRATION_SYSTEM_PROMPT

# Loose faithfulness check needs to tolerate natural-language number
# formatting a correct answer legitimately uses — "1,200" for 1200, "three"
# for 3 — without that becoming a false "hallucination" flag.
_NUMBER_WORDS = {
    '0': 'zero', '1': 'one', '2': 'two', '3': 'three', '4': 'four', '5': 'five',
    '6': 'six', '7': 'seven', '8': 'eight', '9': 'nine', '10': 'ten',
    '11': 'eleven', '12': 'twelve', '13': 'thirteen', '14': 'fourteen',
    '15': 'fifteen', '16': 'sixteen', '17': 'seventeen', '18': 'eighteen',
    '19': 'nineteen', '20': 'twenty',
}


def _fact_present(fact: str, lowered_text: str) -> bool:
    fact_l = fact.lower()
    if fact_l in lowered_text or fact_l in lowered_text.replace(',', ''):
        return True
    word = _NUMBER_WORDS.get(fact_l)
    return bool(word and word in lowered_text)


DEFAULT_MODELS = [
    'qwen3-4b-thinking',
    'google/gemma-4-e4b',
    'nvidia/nemotron-3-nano-4b',
    'qwen/qwen3.5-9b',
    'google/gemma-4-12b',
]

# Each case mirrors a real call site: bot/services.py._build_ai_summary (report
# narration, multi-row) and chat/ai_assistant.py._reply_for (single job card,
# conversational). must_contain is checked case-insensitively as a loose
# faithfulness gate — not a full accuracy grader, just "did it not drop or
# mangle the numbers/identifiers it was given."
CASES = [
    {
        'name': 'job_card_reply',
        'system': CHAT_ASSISTANT_SYSTEM_PROMPT,
        'user': (
            'Question: what is the status of JC-08-26-PP-1595?\n\n'
            'Data:\nJob Card: JC-08-26-PP-1595\nSKU: TAGNYTHREADSHANGTAG72X44\n'
            'PO/WO/PR No: WO-08-2026-06552\nPlate Set No: 11934\nOrder Qty: 139411\n'
            'Status: Pending QC\nCurrent Stage (shop floor): Lamination\n'
            'Machine: GTO 2A\nWastage: 12 sheets (0.4% of total production)\n'
            'Dispatched: 0 of 139411 (0.0%)\n'
        ),
        'must_contain': ['1595', 'pending qc', 'lamination'],
    },
    {
        'name': 'report_narration',
        'system': NARRATION_SYSTEM_PROMPT,
        'user': (
            'Summarize this report:\nReport: Stock Report - Excess Inventory\n'
            'Total records: 3\n'
            'SKU: WIDGET-A, Stock Qty: 500, Bagged: No\n'
            'SKU: WIDGET-B, Stock Qty: 1200, Bagged: Yes\n'
            'SKU: WIDGET-C, Stock Qty: 40, Bagged: No\n'
        ),
        'must_contain': ['3', '500', '1200', '40'],
    },
    {
        # Mirrors a real capped report (bot/report_adapter.py:
        # build_narration_facts caps at 25 rows) — checks the model states
        # the true total (113) rather than the visible sample count (25),
        # and doesn't invent numbers for the 88 rows it never saw.
        'name': 'large_capped_report',
        'system': NARRATION_SYSTEM_PROMPT,
        'user': (
            'Summarize this report:\nReport: Production Pending - Packing\n'
            'Total records: 113\n'
            + '\n'.join(
                f'Job Card: JC-08-26-PP-{1500 + i}, Pending (pcs): {(i + 1) * 137}'
                for i in range(25)
            )
            + '\n... and 88 more row(s) not shown here.\n'
        ),
        'must_contain': ['113', '88'],
    },
    {
        # A report with zero rows — bot/services.py only sends this at all
        # if send_when_empty is True, but when it does, the model must say
        # "no records" plainly rather than speculating about why or padding
        # with invented content.
        'name': 'empty_report',
        'system': NARRATION_SYSTEM_PROMPT,
        'user': 'Summarize this report:\nReport: Wastage Report\nTotal records: 0\n',
        'must_contain': ['0', 'no'],
    },
    {
        # A job card with most optional fields unset — checks the model
        # states "not set"/"not assigned" plainly rather than inventing a
        # plausible-sounding machine name, PO number, or AWC code that was
        # never given to it. This is the single highest-value faithfulness
        # check in the suite: a model that fills in a gap here would
        # fabricate business data in production.
        'name': 'sparse_job_card_reply',
        'system': CHAT_ASSISTANT_SYSTEM_PROMPT,
        'user': (
            'Question: what do you have on JC-08-26-PP-2001?\n\n'
            'Data:\nJob Card: JC-08-26-PP-2001\nSKU: NEWSKU-UNTITLED-0001\n'
            'PO/WO/PR No: Not set\nPlate Set No: Not set\nAWC No: Not set\n'
            'Order Qty: 200\nStatus: Pending Data\n'
            'Current Stage (shop floor): Not Set\nMachine: Not assigned\n'
            'Wastage: 0 sheets (0.0% of total production)\nDispatched: 0 of 200 (0.0%)\n'
        ),
        'must_contain': ['2001', 'not set', 'not assigned'],
    },
]


class Command(BaseCommand):
    help = 'Benchmark candidate LLM models for speed and faithfulness against this app\'s real prompts.'

    def add_arguments(self, parser):
        parser.add_argument('models', nargs='*', default=DEFAULT_MODELS)

    def handle(self, *args, **options):
        results = []
        for model in options['models']:
            for case in CASES:
                self.stdout.write(f'--- {model} / {case["name"]} ---')
                t0 = time.time()
                try:
                    content, usage = call_chat(
                        [
                            {'role': 'system', 'content': case['system']},
                            {'role': 'user', 'content': case['user']},
                        ],
                        model=model,
                        raise_on_error=True,
                        return_usage=True,
                    )
                    elapsed = time.time() - t0
                except Exception as exc:
                    elapsed = time.time() - t0
                    self.stderr.write(f'  FAILED after {elapsed:.1f}s: {exc}')
                    results.append({'model': model, 'case': case['name'], 'elapsed': elapsed, 'ok': False})
                    continue

                lowered = (content or '').lower()
                missing = [f for f in case['must_contain'] if not _fact_present(f, lowered)]
                faithful = not missing
                usage = usage or {}
                reasoning = (usage.get('completion_tokens_details') or {}).get('reasoning_tokens', 0)

                self.stdout.write(f'  {elapsed:.1f}s, {usage.get("completion_tokens", "?")} completion tokens '
                                   f'({reasoning} reasoning), faithful={faithful}'
                                   + (f' MISSING={missing}' if missing else ''))
                self.stdout.write('  ' + (content or '').encode('ascii', 'replace').decode()[:200])

                results.append({
                    'model': model, 'case': case['name'], 'elapsed': elapsed,
                    'ok': True, 'faithful': faithful, 'completion_tokens': usage.get('completion_tokens'),
                    'reasoning_tokens': reasoning,
                })

        self.stdout.write('\n=== Summary ===')
        by_model = {}
        for r in results:
            by_model.setdefault(r['model'], []).append(r)
        for model, rows in by_model.items():
            ok_rows = [r for r in rows if r['ok']]
            if not ok_rows:
                self.stdout.write(f'{model}: all cases FAILED')
                continue
            avg_time = sum(r['elapsed'] for r in ok_rows) / len(ok_rows)
            all_faithful = all(r['faithful'] for r in ok_rows)
            self.stdout.write(f'{model}: avg {avg_time:.1f}s, faithful={all_faithful}, '
                               f'{len(ok_rows)}/{len(rows)} succeeded')
