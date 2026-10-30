import datetime

from django.contrib.auth import get_user_model
from django.core import mail
from django.test import TestCase
from django.urls import reverse
from django.utils import timezone

from core.models import JobCard, Machine, Notification, Production, UserProfile
from planning.models import PlanningJob
from production_plan import engine, services
from production_plan.models import PlanBlock, PlanItem, PlanSettings, ProductionPlan, PlanRunLog
from production_plan.scheduler import run_tick

User = get_user_model()
MON, TUE, WED, FRI, SAT, SUN = (datetime.date(2026, 11, d) for d in (2, 3, 4, 6, 7, 8))


def aware(day, hour, minute=0):
    return timezone.make_aware(datetime.datetime.combine(day, datetime.time(hour, minute)))


class PlanTestBase(TestCase):
    def setUp(self):
        self.cfg = PlanSettings.get()
        self.cfg.shift_capacity = 1000
        self.cfg.friday_day_capacity = 800
        self.cfg.black_machines = 'GTO 1A'
        self.cfg.colour_gto_machines = 'GTO 2C'
        self.cfg.colour_sm_machines = 'SM 74'
        self.cfg.save()
        self.machines = {n: Machine.objects.create(name=n, machine_type='offset_printing', is_active=True) for n in ('GTO 1A', 'GTO 2C', 'SM 74', 'Zongrui')}
        self.counter = 0

    def job(self, remaining, order, colour='1', sheet='10.5*15', new=False, status='released', **extra):
        self.counter += 1
        return PlanningJob.objects.create(
            jc_number=f'JC-T-{self.counter:04d}', sku=f'SKU-{self.counter}', status=status, color_spec=colour, print_sheet_size=sheet,
            planned_total_impressions=remaining, repeat_flag='New' if new else 'Repeat', po_approval_date=order,
            planning_stage='planning_done', **extra,
        )

    def card(self, job, status='in_production'):
        return JobCard.objects.create(
            job_card_no=f'JC-{job.jc_number}', planning_job=job, SKU=job.sku, order_qty=10, ups=1, total_sheet_quantity=100,
            total_impressions_required=1000, total_colors=1, wastage=0, po_date=datetime.date(2026, 10, 1),
            machine_name=self.machines['GTO 1A'], status=status,
        )

    def build(self, morning=TUE, base=None, today=MON):
        return engine.build_plan(morning, self.cfg, base_items=base, today=today)

    def by_slot(self, build, lane=None):
        out = {}
        for it in build.items:
            if lane and it['lane'] != lane:
                continue
            out.setdefault((it['machine_id'], it['plan_date'], it['shift']), []).append(it)
        return out


class CalendarTests(PlanTestBase):
    def test_sunday_is_off_and_friday_day_is_shorter(self):
        self.assertEqual(engine.working_dates_from(SAT, 2, self.cfg), [SAT, datetime.date(2026, 11, 9)])      # Sat, then Mon
        self.assertEqual(engine.next_working_day(SAT, self.cfg), datetime.date(2026, 11, 9))
        self.assertEqual(engine.slot_capacity(FRI, 'D', self.cfg), 800)
        self.assertEqual(engine.slot_capacity(FRI, 'N', self.cfg), 1000)
        self.assertEqual(engine.slot_capacity(SUN, 'D', self.cfg), 0)

    def test_listed_holidays_are_off(self):
        self.cfg.off_dates = '2026-11-03'
        self.cfg.save()
        self.assertEqual(engine.working_dates_from(MON, 2, self.cfg), [MON, WED])


class LaneTests(PlanTestBase):
    def test_one_and_one_plus_one_are_black_and_colour_splits_by_sheet_size(self):
        self.assertEqual(engine.lane_of(self.job(100, MON, colour='1'), self.cfg), 'black')
        self.assertEqual(engine.lane_of(self.job(100, MON, colour='1+1'), self.cfg), 'black')
        self.assertEqual(engine.lane_of(self.job(100, MON, colour='4', sheet='12.5*18'), self.cfg), 'colour_gto')
        self.assertEqual(engine.lane_of(self.job(100, MON, colour='4', sheet='18*25'), self.cfg), 'colour_sm')
        self.assertEqual(engine.lane_of(self.job(100, MON, colour='2+2', sheet='20*29'), self.cfg), 'colour_sm')
        self.assertIsNone(engine.lane_of(self.job(100, MON, colour=''), self.cfg))


class GeneratorRuleTests(PlanTestBase):
    def test_capacity_is_never_exceeded_and_only_lane_machines_are_used(self):
        for i, size in enumerate((100, 900, 300, 2500, 50, 700)):
            self.job(size, datetime.date(2026, 10, 1 + i))
        build = self.build()
        for key, items in self.by_slot(build).items():
            self.assertLessEqual(sum(i['qty'] for i in items), engine.slot_capacity(key[1], key[2], self.cfg))
        self.assertEqual({i['machine_id'] for i in build.items}, {self.machines['GTO 1A'].id})        # Zongrui and the rest never used

    def test_days_take_the_small_end_and_nights_the_large_end(self):
        for i, size in enumerate((50, 100, 200, 300, 400, 900, 800, 700)):
            self.job(size, datetime.date(2026, 10, 1 + i))
        build = self.build()
        day = [i['qty'] for i in build.items if i['shift'] == 'D']
        night = [i['qty'] for i in build.items if i['shift'] == 'N']
        self.assertTrue(day and night)
        self.assertLess(sum(day) / len(day), sum(night) / len(night))                                # small in the day, large at night
        d1 = [i['qty'] for i in sorted(build.items, key=lambda i: i['sequence']) if i['shift'] == 'D' and i['plan_date'] == TUE]
        self.assertEqual(d1, sorted(d1))                                                              # a day shift runs small to large
        n1 = [i['qty'] for i in sorted(build.items, key=lambda i: i['sequence']) if i['shift'] == 'N' and i['plan_date'] == TUE]
        self.assertEqual(n1, sorted(n1, reverse=True))                                                # a night shift runs large to small

    def test_new_jobs_are_never_planned_in_a_night_shift(self):
        for i in range(6):
            self.job(500, datetime.date(2026, 10, 1 + i))
        newbies = [self.job(900, datetime.date(2026, 11, 2), new=True), self.job(950, datetime.date(2026, 11, 2), new=True)]
        build = self.build()
        new_items = [i for i in build.items if i['is_new']]
        self.assertTrue(new_items)
        self.assertTrue(all(i['shift'] == 'D' for i in new_items))
        self.assertEqual({i['planning_job_id'] for i in new_items}, {j.id for j in newbies})

    def test_new_jobs_beyond_the_day_capacity_are_flagged_not_moved_to_night(self):
        for i in range(4):
            self.job(900, datetime.date(2026, 11, 2), new=True)                                       # 3,600 new against 2,000 of day capacity
        build = self.build()
        self.assertTrue(all(i['shift'] == 'D' for i in build.items))
        self.assertTrue(build.summary['overflow_new'])
        self.assertEqual(sum(i['qty'] for i in build.items), 1800)                                    # two fit; a third would be a sliver

    def test_fifo_oldest_jobs_enter_the_plan_when_capacity_is_short(self):
        oldest = self.job(900, datetime.date(2026, 8, 1))
        for i in range(8):
            self.job(900, datetime.date(2026, 10, 10 + i))
        newest = self.job(900, datetime.date(2026, 11, 1))
        planned = {i['planning_job_id'] for i in self.build().items}
        self.assertIn(oldest.id, planned)
        self.assertNotIn(newest.id, planned)                                                           # 10 x 900 does not fit 4,000

    def test_higher_priority_goes_ahead_of_age(self):
        old = self.job(900, datetime.date(2026, 8, 1))
        for i in range(5):
            self.job(900, datetime.date(2026, 8, 2 + i))
        urgent = self.job(900, datetime.date(2026, 11, 1), priority=4)
        planned = {i['planning_job_id'] for i in self.build().items}
        self.assertIn(urgent.id, planned)
        self.assertEqual(len(planned), 4)                                                               # 4,000 capacity: urgent + 3 oldest

    def test_a_job_bigger_than_one_shift_is_cut_into_shift_sized_chunks(self):
        giant = self.job(3500, datetime.date(2026, 9, 1))
        items = [i for i in self.build().items if i['planning_job_id'] == giant.id]
        self.assertTrue(all(i['qty'] <= 1000 and i['is_chunk'] for i in items))
        self.assertEqual(sum(i['qty'] for i in items), 3500)

    def test_printed_work_is_deducted_and_followers_plates_and_holds_are_skipped(self):
        started = self.job(1000, datetime.date(2026, 9, 1), job_process_type='print_and_pack')
        card = self.card(started)
        Production.objects.bulk_create([Production(job_card=card, entry_type='printing', date=MON, shift='A', impressions=600, output_sheets=600)])
        hold = self.job(500, datetime.date(2026, 9, 2), is_on_hold=True)
        plates = self.job(500, datetime.date(2026, 9, 3))
        PlanningJob.objects.filter(pk=plates.pk).update(planning_stage='new_plate_making')
        draft = self.job(500, datetime.date(2026, 9, 4), status='draft')
        build = self.build()
        got = {i['planning_job_id']: sum(x['qty'] for x in build.items if x['planning_job_id'] == i['planning_job_id']) for i in build.items}
        self.assertEqual(got[started.id], 400)                                                          # only what is left
        for skipped in (hold, plates, draft):
            self.assertNotIn(skipped.id, got)

    def test_a_machine_in_preventive_maintenance_gets_no_work_that_shift(self):
        from maintenance.models import PreventiveMaintenancePlan

        for i in range(6):
            self.job(900, datetime.date(2026, 10, 1 + i))
        PreventiveMaintenancePlan.objects.create(machine=self.machines['GTO 1A'], title='PM', interval_value=30, next_due_at=TUE, is_active=True)
        build = self.build()
        self.assertFalse([i for i in build.items if i['plan_date'] == TUE and i['shift'] == 'D'])      # the day shift of the PM day
        self.assertTrue([i for i in build.items if i['plan_date'] == TUE and i['shift'] == 'N'])

    def test_planner_blocks_and_the_friday_day_shift_are_respected(self):
        for i in range(8):
            self.job(900, datetime.date(2026, 10, 1 + i))
        PlanBlock.objects.create(machine=self.machines['GTO 1A'], date=SAT, shift='B', reason='breakdown')
        build = self.build(morning=FRI, today=datetime.date(2026, 11, 5))
        self.assertFalse([i for i in build.items if i['plan_date'] == SAT])
        friday_day = sum(i['qty'] for i in build.items if i['plan_date'] == FRI and i['shift'] == 'D')
        self.assertLessEqual(friday_day, 800)

    def test_lanes_keep_to_their_machines(self):
        black = self.job(500, MON, colour='1+1')
        colour = self.job(500, MON, colour='4', sheet='12.5*18')
        big = self.job(500, MON, colour='4', sheet='20*29')
        build = self.build()
        machine_of = {i['planning_job_id']: i['machine_id'] for i in build.items}
        self.assertEqual(machine_of[black.id], self.machines['GTO 1A'].id)
        self.assertEqual(machine_of[colour.id], self.machines['GTO 2C'].id)
        self.assertEqual(machine_of[big.id], self.machines['SM 74'].id)


class RebalanceTests(PlanTestBase):
    def items(self, build):
        return [dict(i) for i in build.items]

    def test_only_printed_jobs_leave_planner_changes_stay_and_free_capacity_is_refilled(self):
        for i in range(4):
            self.job(400, datetime.date(2026, 9, 1 + i))
        for i in range(10):
            self.job(400, datetime.date(2026, 10, 1 + i))
        base = self.items(self.build())
        planner_item = base[0]
        planner_item['pinned'] = planner_item['manual'] = True
        printed_id = next(i for i in base if i['planning_job_id'] != planner_item['planning_job_id'])['planning_job_id']
        printed_jc = PlanningJob.objects.get(pk=printed_id).jc_number
        PlanningJob.objects.filter(pk=printed_id).update(status='completed')                             # the night finished that job
        second = self.build(base=base)
        before = {i['planning_job_id'] for i in base}
        after = {i['planning_job_id'] for i in second.items}
        self.assertNotIn(printed_id, after)
        kept = [i for i in second.items if i['planning_job_id'] == planner_item['planning_job_id']][0]
        self.assertEqual((kept['machine_id'], kept['plan_date'], kept['shift'], kept['pinned']),
                         (planner_item['machine_id'], planner_item['plan_date'], planner_item['shift'], True))
        self.assertTrue((before - {printed_id}) <= after)                                                 # nothing else was dropped
        self.assertTrue(after - before)                                                                   # freed capacity refilled by the rules
        self.assertIn(printed_jc, second.diff['removed_printed'])

    def test_a_partly_printed_job_is_reduced_to_what_is_left(self):
        job = self.job(1000, datetime.date(2026, 9, 1))
        base = self.items(self.build())
        planned = sum(i['qty'] for i in base if i['planning_job_id'] == job.id)
        card = self.card(job)
        Production.objects.bulk_create([Production(job_card=card, entry_type='printing', date=MON, shift='A', impressions=700, output_sheets=700)])
        after = self.build(base=base)
        self.assertEqual(planned, 1000)
        self.assertEqual(sum(i['qty'] for i in after.items if i['planning_job_id'] == job.id), 300)


class LifecycleAndGateTests(PlanTestBase):
    def setUp(self):
        super().setUp()
        self.planner = User.objects.create_user('plan_user', password='pw12345678', email='planner@example.com', first_name='Pat')
        UserProfile.objects.update_or_create(user=self.planner, defaults={'role': 'planner'})
        self.admin = User.objects.create_superuser('plan_admin', password='pw12345678', email='admin@example.com')
        for i in range(6):
            self.job(800, datetime.date(2026, 10, 1 + i))

    def test_the_evening_draft_notifies_planners_and_creates_the_red_tag_until_approved(self):
        plan = services.build_draft(today=MON, settings_obj=self.cfg)
        self.assertEqual((plan.status, plan.morning_date), ('draft', TUE))
        self.assertEqual(plan.horizon_dates, [TUE.isoformat(), WED.isoformat()])
        note = Notification.objects.get(user=self.planner, event_type='production_plan.draft')
        self.assertIn('approval', note.title)
        self.assertIsNotNone(services.pending_draft())
        self.assertEqual(services.build_draft(today=MON, settings_obj=self.cfg).pk, plan.pk)           # a second run does not duplicate
        services.approve(plan, self.planner)
        self.assertIsNone(services.pending_draft())
        self.assertTrue(Notification.objects.get(pk=note.pk).is_read)

    def test_friday_and_saturday_evenings_plan_for_the_next_working_morning(self):
        self.assertEqual(services.build_draft(today=datetime.date(2026, 11, 7), settings_obj=self.cfg, notify=False).morning_date, datetime.date(2026, 11, 9))

    def test_a_new_draft_supersedes_an_unapproved_one_and_carries_its_items(self):
        first = services.build_draft(today=MON, settings_obj=self.cfg, notify=False)
        item = first.items.first()
        item.pinned = item.manual = True
        item.save()
        second = services.build_draft(today=TUE, settings_obj=self.cfg, notify=False)
        first.refresh_from_db()
        self.assertEqual(first.status, 'superseded')
        self.assertEqual(second.base_plan_id, first.pk)
        if item.plan_date >= WED:
            self.assertTrue(second.items.filter(planning_job_id=item.planning_job_id, pinned=True).exists())

    def test_planners_are_locked_out_of_planning_until_the_plan_is_approved(self):
        self.client.force_login(self.planner)
        plan = services.build_draft(today=MON, settings_obj=self.cfg, notify=False)
        blocked = self.client.get(reverse('planning:home'))
        self.assertEqual(blocked.status_code, 302)
        self.assertEqual(blocked['Location'], reverse('production_plan:board'))
        self.assertEqual(self.client.get(reverse('production_plan:board')).status_code, 200)          # the plan itself stays open
        self.client.logout()
        self.client.force_login(self.admin)
        self.assertNotEqual(self.client.get(reverse('planning:home')).status_code, 302)               # admin is never held
        services.approve(plan, self.admin)                                                             # admin can approve for an absent planner
        self.client.logout()
        self.client.force_login(self.planner)
        self.assertNotEqual(self.client.get(reverse('planning:home')).status_code, 302)               # planning is open again

    def test_the_gate_can_be_switched_off(self):
        services.build_draft(today=MON, settings_obj=self.cfg, notify=False)
        self.cfg.gate_enabled = False
        self.cfg.save()
        self.client.force_login(self.planner)
        self.assertNotEqual(self.client.get(reverse('planning:home')).status_code, 302)


class SchedulerAndEmailTests(PlanTestBase):
    def setUp(self):
        super().setUp()
        self.planner = User.objects.create_user('sched_planner', password='pw12345678', email='planner@example.com')
        UserProfile.objects.update_or_create(user=self.planner, defaults={'role': 'planner'})
        for i in range(6):
            self.job(800, datetime.date(2026, 10, 1 + i))

    def test_the_day_runs_draft_then_rebalance_then_email_once_each(self):
        self.assertEqual(run_tick(aware(MON, 18, 29)), [])
        self.assertEqual(run_tick(aware(MON, 18, 31)), ['draft'])
        self.assertEqual(run_tick(aware(MON, 18, 40)), [])                                              # not twice
        plan = ProductionPlan.objects.get()
        self.assertEqual((plan.morning_date, plan.status), (TUE, 'draft'))

        self.assertEqual(run_tick(aware(TUE, 8, 31)), ['rebalance'])
        plan.refresh_from_db()
        self.assertIsNotNone(plan.rebalanced_at)
        self.assertEqual(mail.outbox, [])                                                               # nothing before 08:45

        self.assertEqual(run_tick(aware(TUE, 8, 46)), ['email'])
        self.assertEqual(len(mail.outbox), 1)
        message = mail.outbox[0]
        self.assertEqual(message.to, ['planner@example.com'])
        self.assertEqual(message.cc, ['printing@utopia.pk'])
        self.assertIn('not re-approved', message.subject)                                               # nobody approved it
        self.assertEqual(sorted(a[0].rsplit('.', 1)[1] for a in message.attachments), ['pdf', 'xlsx'])
        self.assertIn('NOT RE-APPROVED', message.alternatives[0][0])
        self.assertEqual(run_tick(aware(TUE, 8, 50)), [])                                               # sent once
        plan.refresh_from_db()
        self.assertEqual(plan.email_status, 'sent_unapproved')

    def test_an_approved_plan_is_emailed_as_approved(self):
        run_tick(aware(MON, 18, 31))
        plan = ProductionPlan.objects.get()
        services.approve(plan, self.planner)
        run_tick(aware(TUE, 8, 31))
        run_tick(aware(TUE, 8, 46))
        self.assertNotIn('not re-approved', mail.outbox[0].subject)
        self.assertIn('Approved by', mail.outbox[0].alternatives[0][0])

    def test_sunday_has_no_morning_run_and_saturday_evening_plans_monday(self):
        self.assertEqual(run_tick(aware(SUN, 18, 31)), [])
        self.assertEqual(run_tick(aware(SUN, 8, 31)), [])
        self.assertEqual(run_tick(aware(SAT, 18, 31)), ['draft'])
        self.assertEqual(ProductionPlan.objects.get().morning_date, datetime.date(2026, 11, 9))

    def test_a_missed_draft_still_runs_later_in_the_evening_and_the_switch_stops_everything(self):
        self.assertEqual(run_tick(aware(MON, 21, 15)), ['draft'])
        self.cfg.scheduler_enabled = False
        self.cfg.save()
        self.assertEqual(run_tick(aware(TUE, 8, 31)), [])

    def test_a_failed_email_is_retried_then_recorded(self):
        run_tick(aware(MON, 18, 31))
        run_tick(aware(TUE, 8, 31))
        self.cfg.email_cc = ''
        self.cfg.email_to_roles = 'nobody'
        self.cfg.save()                                                                                  # no recipients: the send fails
        self.assertEqual(run_tick(aware(TUE, 8, 46)), ['email'])
        row = PlanRunLog.objects.get(task='email')
        self.assertEqual(row.status, 'retry')
        self.assertEqual(ProductionPlan.objects.get().email_status, 'failed')
        self.cfg.email_cc = 'printing@utopia.pk'
        self.cfg.save()
        self.assertEqual(run_tick(aware(TUE, 8, 48)), [])                                                # waits the retry interval
        self.assertEqual(run_tick(aware(TUE, 8, 52)), ['email'])
        self.assertEqual(len(mail.outbox), 1)


class ViewTests(PlanTestBase):
    def setUp(self):
        super().setUp()
        self.planner = User.objects.create_user('view_planner', password='pw12345678')
        UserProfile.objects.update_or_create(user=self.planner, defaults={'role': 'planner'})
        self.viewer = User.objects.create_user('view_viewer', password='pw12345678')
        UserProfile.objects.update_or_create(user=self.viewer, defaults={'role': 'viewer'})
        self.admin = User.objects.create_superuser('view_admin', password='pw12345678')
        self.jobs = [self.job(700, datetime.date(2026, 10, 1 + i)) for i in range(5)]
        self.extra = self.job(300, datetime.date(2026, 11, 1))
        self.plan = services.build_draft(today=MON, settings_obj=self.cfg, notify=False)

    def test_board_exports_and_sharing(self):
        self.client.force_login(self.planner)
        html = self.client.get(reverse('production_plan:board')).content.decode()
        for expected in ('Approve the plan', 'Waiting for approval', 'GTO 1A', 'Add a job to the plan'):
            self.assertIn(expected, html)
        for tab in ('actual', 'history'):
            self.assertEqual(self.client.get(reverse('production_plan:board'), {'tab': tab}).status_code, 200)
        xlsx = self.client.get(reverse('production_plan:export', args=[self.plan.id, 'xlsx']))
        pdf = self.client.get(reverse('production_plan:export', args=[self.plan.id, 'pdf']))
        self.assertEqual((xlsx.status_code, pdf.status_code), (200, 200))
        self.assertTrue(pdf.content.startswith(b'%PDF'))
        text = self.client.get(reverse('production_plan:share_text', args=[self.plan.id])).json()['text']
        self.assertIn('*PRODUCTION PLAN*', text)
        self.assertIn('NOT APPROVED', text)
        image = self.client.get(reverse('production_plan:share_image', args=[self.plan.id]))
        self.assertEqual((image.status_code, image['Content-Type']), (200, 'image/png'))
        self.assertEqual(image.content[:8], b'\x89PNG\r\n\x1a\n')

    def test_only_planners_and_admins_edit_everyone_else_reads(self):
        self.client.force_login(self.viewer)
        self.assertEqual(self.client.get(reverse('production_plan:board')).status_code, 200)
        html = self.client.get(reverse('production_plan:board')).content.decode()
        self.assertNotIn('Approve the plan', html)
        self.assertEqual(self.client.post(reverse('production_plan:approve', args=[self.plan.id])).status_code, 403)
        item = self.plan.items.first()
        self.assertEqual(self.client.post(reverse('production_plan:item_remove', args=[self.plan.id, item.id])).status_code, 403)
        self.plan.refresh_from_db()
        self.assertEqual(self.plan.status, 'draft')

    def test_planner_approves_and_the_red_tag_clears(self):
        self.client.force_login(self.planner)
        self.assertEqual(self.client.get(reverse('production_plan:board')).context['nav']['plan_pending_count'], 1)     # the red tag
        self.client.force_login(self.viewer)
        self.assertEqual(self.client.get(reverse('production_plan:board')).context['nav']['plan_pending_count'], 0)      # not for others
        self.client.force_login(self.planner)
        self.client.post(reverse('production_plan:approve', args=[self.plan.id]))
        self.plan.refresh_from_db()
        self.assertEqual((self.plan.status, self.plan.approved_by), ('approved', self.planner))
        self.assertEqual(self.client.get(reverse('production_plan:board')).context['nav']['plan_pending_count'], 0)

    def test_planner_edits_remove_move_pin_priority_and_blocks(self):
        self.client.force_login(self.planner)
        item = self.plan.items.filter(shift='D').first()
        day_slot = f"{item.plan_date.isoformat()}|N" if item.plan_date else ''
        self.client.post(reverse('production_plan:item_pin', args=[self.plan.id, item.id]))
        item.refresh_from_db()
        self.assertTrue(item.pinned)
        self.client.post(reverse('production_plan:item_move', args=[self.plan.id, item.id]),
                         {'machine_id': item.machine_id, 'slot': day_slot, 'qty': item.qty})
        self.assertFalse(PlanItem.objects.filter(pk=item.pk).exists() and PlanItem.objects.get(pk=item.pk).shift == 'D' and item.qty == 0)
        moved = self.plan.items.filter(planning_job=item.planning_job, shift='N', manual=True).first()
        self.assertIsNotNone(moved)
        self.assertTrue(moved.pinned)
        self.client.post(reverse('production_plan:item_remove', args=[self.plan.id, moved.id]))
        self.assertFalse(PlanItem.objects.filter(pk=moved.pk).exists())
        self.client.post(reverse('production_plan:set_priority', args=[self.plan.id]), {'job_id': self.extra.id, 'priority': 4})
        self.extra.refresh_from_db()
        self.assertEqual(self.extra.priority, 4)
        self.client.post(reverse('production_plan:block_add', args=[self.plan.id]),
                         {'machine_id': self.machines['GTO 1A'].id, 'date': WED.isoformat(), 'shift': 'D', 'reason': 'service'})
        self.assertTrue(PlanBlock.objects.filter(date=WED, shift='D').exists())
        self.assertGreaterEqual(self.plan.events.count(), 5)                                           # every change is logged

    def test_adding_a_job_outside_fifo_needs_a_reason_and_a_new_job_cannot_go_to_night(self):
        self.client.force_login(self.planner)
        for i in range(12):
            self.job(900, datetime.date(2026, 8, 1 + i))                                               # a long queue ahead of the extra job
        late = self.job(100, datetime.date(2026, 11, 2))
        slot = f'{TUE.isoformat()}|D'
        data = {'lane': 'black', 'job_id': late.id, 'machine_id': self.machines['GTO 1A'].id, 'slot': slot, 'qty': 100}
        self.client.post(reverse('production_plan:item_add', args=[self.plan.id]), data)
        self.assertFalse(self.plan.items.filter(planning_job=late, manual=True).exists())               # refused: outside FIFO, no reason
        self.client.post(reverse('production_plan:item_add', args=[self.plan.id]), {**data, 'fifo_reason': 'customer is waiting'})
        added = self.plan.items.get(planning_job=late, manual=True)
        self.assertEqual((added.fifo_reason, added.pinned), ('customer is waiting', True))
        newbie = self.job(100, datetime.date(2026, 11, 2), new=True)
        self.client.post(reverse('production_plan:item_add', args=[self.plan.id]),
                         {'lane': 'black', 'job_id': newbie.id, 'machine_id': self.machines['GTO 1A'].id, 'slot': f'{TUE.isoformat()}|N', 'qty': 100})
        self.assertFalse(self.plan.items.filter(planning_job=newbie).exists())

    def test_job_search_and_settings_page(self):
        self.client.force_login(self.planner)
        found = self.client.get(reverse('production_plan:job_search'), {'q': self.extra.jc_number[-4:]}).json()['results']
        self.assertEqual([r['jc'] for r in found], [self.extra.jc_number])
        self.assertEqual(self.client.get(reverse('production_plan:settings')).status_code, 403)         # admin only
        self.client.force_login(self.admin)
        page = self.client.get(reverse('production_plan:settings'))
        self.assertEqual(page.status_code, 200)

    def test_plan_vs_actual_counts_entries_by_machine_and_shift(self):
        job = self.jobs[0]
        card = self.card(job)
        Production.objects.bulk_create([Production(job_card=card, entry_type='printing', date=TUE, shift='A',
                                                   machine=self.machines['GTO 1A'], impressions=500, output_sheets=500)])
        rows = services.plan_vs_actual(self.plan, self.cfg, as_of=aware(WED, 9))
        day = [r for r in rows if r['date'] == TUE and r['shift'] == 'D' and r['machine'] == 'GTO 1A'][0]
        self.assertEqual((day['actual'], day['started']), (500, True))
        future = [r for r in rows if r['date'] == WED and r['shift'] == 'N'][0]
        self.assertFalse(future['started'])
