import time

from django.core.management.base import BaseCommand

from core.llm.client import call_chat
from core.llm.prompts import NARRATION_SYSTEM_PROMPT
from core.models import JobCard


class Command(BaseCommand):
    help = 'Proof of concept: fetch one JobCard via ORM, ask the local LLM to summarize it.'

    def add_arguments(self, parser):
        parser.add_argument('job_card_no', nargs='?', default=None)

    def handle(self, *args, **options):
        jc = (
            JobCard.objects.filter(job_card_no=options['job_card_no']).first()
            if options['job_card_no']
            else JobCard.objects.order_by('-id').first()
        )
        if not jc:
            self.stderr.write('No matching job card found.')
            return

        facts = (
            f"Job Card: {jc.job_card_no}\n"
            f"SKU: {jc.SKU}\n"
            f"Order Qty: {jc.order_qty}\n"
            f"Status: {jc.workflow_status_label}\n"
            f"Machine: {jc.machine_name.name if jc.machine_name else 'Not assigned'}\n"
        )
        self.stdout.write(f'--- Facts sent to LLM ---\n{facts}')

        t0 = time.time()
        result = call_chat([
            {'role': 'system', 'content': NARRATION_SYSTEM_PROMPT},
            {'role': 'user', 'content': f'Summarize this job card:\n{facts}'},
        ], raise_on_error=True)
        self.stdout.write(f'--- LLM summary ({time.time() - t0:.1f}s) ---\n{result}')
