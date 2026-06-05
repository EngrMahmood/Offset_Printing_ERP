from datetime import date

from django.core.exceptions import ValidationError
from django.test import TestCase
from django.contrib.auth import get_user_model

from .jc_numbering import allocate_next_jc_number
from .models import JobCard, Machine, Production, Dispatch, SequenceCounter


class DispatchValidationTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(username='tester', password='testpass')
        self.machine = Machine.objects.create(name='Test Machine')

    def test_print_job_dispatch_within_produced_pieces_succeeds(self):
        job_card = JobCard.objects.create(
            job_card_no='JC-01-26-1071',
            order_qty=100,
            ups=10,
            is_print_job=True,
            total_impressions_required=100,
        )
        Production.objects.create(
            job_card=job_card,
            date='2026-01-01',
            shift='A',
            machine=self.machine,
            output_sheets=10,
            waste_sheets=0,
            impressions=100,
            planned_time=60,
            run_time=60,
        )

        dispatch = Dispatch(
            job_card=job_card,
            dispatch_date='2026-01-02',
            dispatch_qty=50,
            created_by=self.user,
        )
        dispatch.save()

        self.assertEqual(job_card.total_dispatch, 50)

    def test_print_job_dispatch_exceeding_production_raises(self):
        job_card = JobCard.objects.create(
            job_card_no='JC-02-26-0001',
            order_qty=100,
            ups=10,
            is_print_job=True,
            total_impressions_required=50,
        )
        Production.objects.create(
            job_card=job_card,
            date='2026-01-01',
            shift='A',
            machine=self.machine,
            output_sheets=5,
            waste_sheets=0,
            impressions=50,
            planned_time=60,
            run_time=60,
        )

        dispatch = Dispatch(
            job_card=job_card,
            dispatch_date='2026-01-02',
            dispatch_qty=60,
            created_by=self.user,
        )
        with self.assertRaises(ValidationError):
            dispatch.save()

    def test_allocate_next_jc_number_always_includes_pp(self):
        SequenceCounter.objects.all().delete()
        job_card_no = allocate_next_jc_number(date(2026, 6, 26))
        self.assertRegex(job_card_no, r'^JC-06-26-PP-\d{4}$')

