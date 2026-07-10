from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.db.models import Q
from django.urls import reverse, reverse_lazy
from django.utils import timezone
from django.contrib import messages
from django.views.generic import ListView, DetailView, CreateView, TemplateView

from planning.models import PlanningJob
from planning.services import normalize_awc_no, normalize_die_cutting, _user_is_admin
from .forms import PlateRequestForm
from .models import PlateRequest
from .services import (
    PLATE_REQUEST_TYPE_FILTERS,
    build_plate_request_type_counts,
    build_type_filter_sidebar,
    build_vendor_filter_options,
    bulk_cancel_stale_open_plate_requests,
    filter_plate_requests_by_type,
    plate_request_active_queryset,
    stale_open_plate_requests_for_cleanup_queryset,
)
from core.models import Vendor


class GraphicsDesignerAccessMixin(UserPassesTestMixin):
    def test_func(self):
        profile = getattr(self.request.user, 'profile', None)
        return bool(profile and profile.can_view_plate_queue())


class PlateDashboardView(LoginRequiredMixin, GraphicsDesignerAccessMixin, TemplateView):
    template_name = 'printing_plates/plate_dashboard.html'

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)

        # 1. Plate Request Counts
        req_qs = plate_request_active_queryset()
        context.update(build_plate_request_type_counts(req_qs))
        context['req_all_count'] = context['all_count']
        context['req_repeat_count'] = context['repeat_count']
        context['req_new_artwork_count'] = context['new_artwork_count']
        context['req_replacement_count'] = context['replacement_count']

        # 2. Plate Sent Counts (vendors from master + used values)
        sent_qs = PlateRequest.objects.filter(
            status=PlateRequest.STATUS_SENT,
            planning_job__isnull=False,
            planning_job__planning_stage__in=['new_plate_making', 'repeat_plate_making']
        )
        context['sent_all_count'] = sent_qs.count()
        context['sent_vendor_options'] = build_vendor_filter_options(sent_qs)

        # 3. Plate Received Counts
        rec_qs = PlateRequest.objects.filter(
            status=PlateRequest.STATUS_RECEIVED,
            planning_job__isnull=False,
            planning_job__planning_stage__in=['new_plate_making', 'repeat_plate_making']
        )
        context['rec_all_count'] = rec_qs.count()
        context['rec_empty_count'] = rec_qs.filter(vendor='').count()
        context['rec_vendor_options'] = build_vendor_filter_options(rec_qs)

        return context


class PlateRequestListView(LoginRequiredMixin, GraphicsDesignerAccessMixin, ListView):
    model = PlateRequest
    template_name = 'printing_plates/plate_request_list.html'
    context_object_name = 'plate_requests'
    paginate_by = 50

    def get_queryset(self):
        queryset = plate_request_active_queryset().select_related(
            'planning_job', 'job_card', 'sku_recipe', 'machine', 'department', 'requested_by'
        ).order_by('-requested_at', '-created_at')

        self.request_type = self.request.GET.get('type', '').strip().lower()
        queryset = filter_plate_requests_by_type(queryset, self.request_type)

        # General search filter
        q = self.request.GET.get('q', '').strip()
        if q:
            queryset = queryset.filter(
                Q(job_card__job_card_no__icontains=q) |
                Q(planning_job__jc_number__icontains=q) |
                Q(planning_job__job_name__icontains=q) |
                Q(planning_job__sku__icontains=q) |
                Q(sku_recipe__sku__icontains=q) |
                Q(plate_color__icontains=q)
            )

        return queryset

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)

        base_qs = plate_request_active_queryset()
        counts = build_plate_request_type_counts(base_qs)
        context.update(counts)
        context['type_filters'] = build_type_filter_sidebar(counts)

        context['all_count'] = counts['all_count']
        context['request_type'] = self.request_type
        context['q'] = self.request.GET.get('q', '')
        context['list_count'] = self.get_queryset().count()
        context['stale_cleanup_count'] = stale_open_plate_requests_for_cleanup_queryset().count()
        context['can_admin_stale_cleanup'] = _user_is_admin(self.request.user)
        return context


class PlateQueueView(LoginRequiredMixin, GraphicsDesignerAccessMixin, ListView):
    """Active plate work: open plate requests only (same source of truth as Plate Requests)."""

    model = PlateRequest
    template_name = 'printing_plates/plate_queue.html'
    context_object_name = 'plate_requests'
    paginate_by = 50

    def get_queryset(self):
        queryset = plate_request_active_queryset().filter(
            status__in=[
                PlateRequest.STATUS_DRAFT,
                PlateRequest.STATUS_SENT,
                PlateRequest.STATUS_RECEIVED,
            ],
        ).select_related(
            'planning_job', 'job_card', 'sku_recipe', 'machine', 'department', 'requested_by'
        ).order_by('-requested_at', '-created_at')

        q = self.request.GET.get('q', '').strip()
        if q:
            queryset = queryset.filter(
                Q(job_card__job_card_no__icontains=q)
                | Q(planning_job__jc_number__icontains=q)
                | Q(planning_job__job_name__icontains=q)
                | Q(planning_job__sku__icontains=q)
                | Q(sku_recipe__sku__icontains=q)
                | Q(awc_no__icontains=q)
                | Q(set_no__icontains=q)
                | Q(new_set_no__icontains=q)
            )

        self.request_type = self.request.GET.get('type', '').strip().lower()
        queryset = filter_plate_requests_by_type(queryset, self.request_type)

        self.status_filter = self.request.GET.get('status', '').strip()
        if self.status_filter in {
            PlateRequest.STATUS_DRAFT,
            PlateRequest.STATUS_SENT,
            PlateRequest.STATUS_RECEIVED,
        }:
            queryset = queryset.filter(status=self.status_filter)

        return queryset

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        base_qs = plate_request_active_queryset().filter(
            status__in=[
                PlateRequest.STATUS_DRAFT,
                PlateRequest.STATUS_SENT,
                PlateRequest.STATUS_RECEIVED,
            ],
        )
        context['q'] = self.request.GET.get('q', '')
        context['request_type'] = getattr(self, 'request_type', '')
        context['status_filter'] = getattr(self, 'status_filter', '')
        context['queue_count'] = self.get_queryset().count()
        counts = build_plate_request_type_counts(base_qs)
        context.update(counts)
        context['type_filters'] = build_type_filter_sidebar(
            counts,
            filters=[item for item in PLATE_REQUEST_TYPE_FILTERS if item['key'] in {'', 'repeat', 'new_artwork', 'replacement'}],
        )
        context['draft_count'] = base_qs.filter(status=PlateRequest.STATUS_DRAFT).count()
        context['sent_count'] = base_qs.filter(status=PlateRequest.STATUS_SENT).count()
        context['received_count'] = base_qs.filter(status=PlateRequest.STATUS_RECEIVED).count()
        return context





class PlateSentListView(LoginRequiredMixin, GraphicsDesignerAccessMixin, ListView):
    model = PlateRequest
    template_name = 'printing_plates/plate_sent_list.html'
    context_object_name = 'plate_requests'
    paginate_by = 50

    def get_queryset(self):
        queryset = PlateRequest.objects.filter(
            status=PlateRequest.STATUS_SENT,
            planning_job__isnull=False,
            planning_job__planning_stage__in=['new_plate_making', 'repeat_plate_making']
        ).select_related(
            'planning_job', 'job_card', 'sku_recipe', 'machine', 'department', 'requested_by', 'sent_by'
        ).order_by('-sent_at', '-requested_at')

        # Vendor filter from sidebar (soft-coded master + used vendors)
        self.vendor_filter = self.request.GET.get('vendor', '').strip()
        if self.vendor_filter:
            queryset = queryset.filter(vendor__iexact=self.vendor_filter)

        # Search filter
        q = self.request.GET.get('q', '').strip()
        if q:
            queryset = queryset.filter(
                Q(job_card__job_card_no__icontains=q) |
                Q(planning_job__jc_number__icontains=q) |
                Q(planning_job__job_name__icontains=q) |
                Q(planning_job__sku__icontains=q) |
                Q(sku_recipe__sku__icontains=q) |
                Q(plate_color__icontains=q)
            )

        return queryset

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        base_qs = PlateRequest.objects.filter(
            status=PlateRequest.STATUS_SENT,
            planning_job__isnull=False,
            planning_job__planning_stage__in=['new_plate_making', 'repeat_plate_making']
        )
        context['vendor_options'] = build_vendor_filter_options(base_qs)
        context['all_count'] = base_qs.count()
        context['vendor_filter'] = self.vendor_filter
        context['q'] = self.request.GET.get('q', '')
        return context


class PlateReceivedListView(LoginRequiredMixin, GraphicsDesignerAccessMixin, ListView):
    model = PlateRequest
    template_name = 'printing_plates/plate_received_list.html'
    context_object_name = 'plate_requests'
    paginate_by = 50

    def get_queryset(self):
        queryset = PlateRequest.objects.filter(
            status=PlateRequest.STATUS_RECEIVED,
            planning_job__isnull=False,
            planning_job__planning_stage__in=['new_plate_making', 'repeat_plate_making']
        ).select_related(
            'planning_job', 'job_card', 'sku_recipe', 'machine', 'department', 'requested_by', 'received_by'
        ).order_by('-received_at', '-requested_at')

        # Vendor filter from sidebar (soft-coded master + used vendors)
        self.vendor_filter = self.request.GET.get('vendor', '').strip()
        if self.vendor_filter == 'empty':
            queryset = queryset.filter(vendor='')
        elif self.vendor_filter:
            queryset = queryset.filter(vendor__iexact=self.vendor_filter)

        # Search filter
        q = self.request.GET.get('q', '').strip()
        if q:
            queryset = queryset.filter(
                Q(job_card__job_card_no__icontains=q) |
                Q(planning_job__jc_number__icontains=q) |
                Q(planning_job__job_name__icontains=q) |
                Q(planning_job__sku__icontains=q) |
                Q(sku_recipe__sku__icontains=q) |
                Q(plate_color__icontains=q)
            )

        return queryset

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        base_qs = PlateRequest.objects.filter(
            status=PlateRequest.STATUS_RECEIVED,
            planning_job__isnull=False,
            planning_job__planning_stage__in=['new_plate_making', 'repeat_plate_making']
        )
        context['vendor_options'] = build_vendor_filter_options(base_qs)
        context['empty_count'] = base_qs.filter(vendor='').count()
        context['all_count'] = base_qs.count()
        context['vendor_filter'] = self.vendor_filter
        context['q'] = self.request.GET.get('q', '')
        return context


class PlateRequestDetailView(LoginRequiredMixin, GraphicsDesignerAccessMixin, DetailView):
    model = PlateRequest
    template_name = 'printing_plates/plate_request_detail.html'
    context_object_name = 'plate_request'

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['vendors'] = Vendor.objects.filter(is_active=True).order_by('name')
        from printing_plates.constants import PLATE_INK_OPTIONS
        context['plate_ink_options'] = PLATE_INK_OPTIONS

        # Resolve the best available master recipe for this plate request's SKU,
        # falling back through: approved > reviewed > pending_review > draft.
        # This ensures the designer sees layout specs even when the planning job
        # fields were not yet populated at the time the plate request was created.
        plate_request = self.object
        sku = (
            (plate_request.sku_recipe.sku if plate_request.sku_recipe else None)
            or (plate_request.planning_job.sku if plate_request.planning_job else None)
            or ''
        ).strip()

        master_recipe = plate_request.sku_recipe
        if sku:
            from planning.services import (
                ensure_sku_recipe_for_planning_job,
                get_best_sku_recipe_for_sku,
                sync_planning_job_fields_to_sku_recipe,
            )
            master_recipe = master_recipe or get_best_sku_recipe_for_sku(sku)
            if plate_request.planning_job and not master_recipe:
                master_recipe = ensure_sku_recipe_for_planning_job(
                    plate_request.planning_job,
                    actor=self.request.user,
                )
            if plate_request.planning_job and master_recipe:
                sync_planning_job_fields_to_sku_recipe(plate_request.planning_job, master_recipe)

        context['master_recipe'] = master_recipe

        from core.print_colors import get_print_color_choices, resolve_print_color_name

        from printing_plates.constants import is_plate_ink_spec

        current_print_color = ''
        for candidate in (
            (plate_request.planning_job.color_spec if plate_request.planning_job else ''),
            (master_recipe.color_spec if master_recipe else ''),
        ):
            candidate = (candidate or '').strip()
            if candidate and not is_plate_ink_spec(candidate):
                current_print_color = candidate
                break
        context['current_print_color'] = current_print_color
        context['print_color_resolved'] = resolve_print_color_name(current_print_color) or current_print_color
        context['print_color_missing'] = not bool(resolve_print_color_name(current_print_color))
        context['print_color_choices'] = get_print_color_choices(include_legacy=current_print_color)

        from printing_plates.plate_set_helpers import build_plate_set_suggestion

        suggestion = build_plate_set_suggestion(plate_request)
        context['plate_set_suggestion'] = suggestion
        context['suggested_sets_required'] = plate_request.sets_required or suggestion['sets_required']
        context['suggested_plate_quantity'] = (
            plate_request.plate_quantity
            if plate_request.plate_quantity is not None
            else suggestion['plate_quantity']
        )
        context['can_correct_designer_layout'] = (
            not master_recipe or (master_recipe.master_data_status or '') != 'approved'
        )
        context['master_sku_locked'] = bool(
            master_recipe and (master_recipe.master_data_status or '') == 'approved'
        )
        return context


from django.views import View
from django.shortcuts import get_object_or_404, redirect

class PlateRequestActionView(LoginRequiredMixin, GraphicsDesignerAccessMixin, View):
    def post(self, request, pk):
        plate_request = get_object_or_404(PlateRequest, pk=pk)
        action = request.POST.get('action')
        
        # Check permissions for designer/admin/manager
        profile = getattr(request.user, 'profile', None)
        role = profile.role if profile else ''
        if role not in ('admin', 'manager', 'graphics_designer') and not request.user.is_superuser:
            messages.error(request, "Only Graphics Designers, Managers, or Admins can perform these actions.")
            return redirect('printing_plates:request_detail', pk=pk)
            
        if action == 'send_to_vendor':
            vendor = request.POST.get('vendor', '').strip()
            set_no = request.POST.get('set_no', '').strip()
            new_set_no = request.POST.get('new_set_no', '').strip()
            awc_no = normalize_awc_no(request.POST.get('awc_no', ''))
            plate_color = request.POST.get('plate_color', '').strip()
            print_color = request.POST.get('print_color', '').strip()
            plate_quantity = request.POST.get('plate_quantity', '').strip()
            sets_required_raw = request.POST.get('sets_required', '').strip()
            remarks = request.POST.get('remarks', '').strip()
            
            if not vendor:
                messages.error(request, "Vendor is required.")
                return redirect('printing_plates:request_detail', pk=pk)
            if not plate_color:
                messages.error(request, "Plate inks are required (Cyan, Magenta, Yellow, Black, Special 1/2).")
                return redirect('printing_plates:request_detail', pk=pk)

            from core.print_colors import (
                apply_print_color_to_planning_job,
                apply_print_color_to_sku_recipe,
                resolve_print_color_name,
            )
            from planning.services import (
                apply_designer_layout_to_sku_recipe,
                designer_layout_missing_fields,
                designer_layout_validation_errors,
                ensure_sku_recipe_for_planning_job,
                get_awc_conflict_message,
                normalize_sheet_size,
                resolve_designer_layout_values,
                sync_planning_job_fields_to_sku_recipe,
                _missing_required_master_fields,
            )
            from printing_plates.plate_set_helpers import build_plate_set_suggestion

            planning_job = plate_request.planning_job
            recipe = plate_request.sku_recipe
            if planning_job and not recipe:
                recipe = ensure_sku_recipe_for_planning_job(planning_job, actor=request.user)
                if recipe:
                    plate_request.sku_recipe = recipe
                    plate_request.save(update_fields=['sku_recipe'])

            sku_for_awc = (
                (recipe.sku if recipe else '')
                or (planning_job.sku if planning_job else '')
                or (plate_request.job_card.SKU if plate_request.job_card_id else '')
                or ''
            ).strip()
            awc_conflict = get_awc_conflict_message(
                awc_no,
                sku=sku_for_awc,
                exclude_recipe_id=recipe.pk if recipe and recipe.pk else None,
                exclude_plate_request_id=plate_request.pk,
            )
            if awc_conflict:
                messages.error(request, awc_conflict)
                return redirect('printing_plates:request_detail', pk=pk)

            from printing_plates.constants import is_plate_ink_spec

            # Never treat plate-ink chips as production print color.
            existing_print_color = ''
            for candidate in (
                (planning_job.color_spec if planning_job else ''),
                (recipe.color_spec if recipe else ''),
            ):
                candidate = (candidate or '').strip()
                if candidate and not is_plate_ink_spec(candidate):
                    existing_print_color = candidate
                    break

            resolved_print_color = resolve_print_color_name(print_color) or resolve_print_color_name(existing_print_color)
            if not resolved_print_color:
                messages.error(
                    request,
                    "Print Color is required (production pattern: 1, 2, 4, 1+1, …). Select from the master list.",
                )
                return redirect('printing_plates:request_detail', pk=pk)

            suggestion = build_plate_set_suggestion(plate_request, plate_color=plate_color)
            try:
                sets_required = int(sets_required_raw) if sets_required_raw else int(suggestion['sets_required'])
            except (TypeError, ValueError):
                messages.error(request, "Sets required must be a valid integer.")
                return redirect('printing_plates:request_detail', pk=pk)
            if sets_required < 1:
                messages.error(request, "Sets required must be at least 1.")
                return redirect('printing_plates:request_detail', pk=pk)

            if plate_quantity:
                try:
                    plate_quantity_value = int(plate_quantity)
                except ValueError:
                    messages.error(request, "Plate Quantity must be a valid integer.")
                    return redirect('printing_plates:request_detail', pk=pk)
            else:
                plate_quantity_value = suggestion['plate_quantity']
            if not plate_quantity_value or plate_quantity_value < 1:
                messages.error(request, "Plate Quantity (total plates) is required.")
                return redirect('printing_plates:request_detail', pk=pk)

            if not plate_request.machine_id and suggestion.get('machine'):
                plate_request.machine = suggestion['machine']

            plate_request.vendor = vendor
            plate_request.set_no = set_no
            plate_request.new_set_no = new_set_no
            plate_request.awc_no = awc_no
            plate_request.plate_color = plate_color
            plate_request.sets_required = sets_required
            plate_request.plate_quantity = plate_quantity_value
            plate_request.remarks = remarks

            if planning_job:
                planning_job.remarks = remarks
                if resolved_print_color:
                    apply_print_color_to_planning_job(planning_job, resolved_print_color)
                    planning_job.save()
                else:
                    planning_job.save(update_fields=['remarks'])

            if recipe:
                recipe.remarks = remarks
                recipe.notes = remarks
                if resolved_print_color:
                    apply_print_color_to_sku_recipe(recipe, resolved_print_color)
                    recipe.save()
                else:
                    recipe.save(update_fields=['remarks', 'notes'])

            # Always block Send to Vendor when designer fields are missing (except remarks).
            master_locked = bool(recipe and recipe.master_data_status == 'approved')
            designer_post = {}
            if not master_locked:
                designer_post = {
                    'size_w_mm': request.POST.get('size_w_mm', '').strip(),
                    'size_h_mm': request.POST.get('size_h_mm', '').strip(),
                    'print_sheet_size': request.POST.get('print_sheet_size', '').strip(),
                    'ups': request.POST.get('ups', '').strip(),
                    'purchase_sheet_size': request.POST.get('purchase_sheet_size', '').strip(),
                    'purchase_sheet_ups': request.POST.get('purchase_sheet_ups', '').strip(),
                    'die_cutting': request.POST.get('die_cutting', '').strip(),
                }

            layout_values = resolve_designer_layout_values(
                {
                    **designer_post,
                    'awc_no': awc_no,
                    'set_no': set_no,
                    'new_set_no': new_set_no,
                },
                recipe=recipe,
                planning_job=planning_job,
                plate_request=plate_request,
            )
            missing_designer = designer_layout_missing_fields(layout_values)
            layout_errors = designer_layout_validation_errors(layout_values)
            if missing_designer or layout_errors:
                messages.error(
                    request,
                    "Cannot send to vendor. Missing required fields (remarks optional): "
                    + ", ".join(missing_designer + layout_errors)
                    + ".",
                )
                return redirect('printing_plates:request_detail', pk=pk)

            if planning_job and not recipe:
                recipe = ensure_sku_recipe_for_planning_job(planning_job, actor=request.user)
                if recipe:
                    plate_request.sku_recipe = recipe
                    plate_request.save(update_fields=['sku_recipe'])

            if recipe and not master_locked:
                try:
                    apply_designer_layout_to_sku_recipe(
                        planning_job,
                        recipe,
                        {
                            **layout_values,
                            'plate_color': plate_color,
                        },
                    )
                except ValueError as exc:
                    messages.error(request, str(exc))
                    return redirect('printing_plates:request_detail', pk=pk)
                recipe.refresh_from_db()

            if planning_job:
                try:
                    planning_job.size_w_mm = int(round(float(layout_values['size_w_mm'])))
                    planning_job.size_h_mm = int(round(float(layout_values['size_h_mm'])))
                    planning_job.ups = int(round(float(layout_values['ups'])))
                    planning_job.purchase_sheet_ups = int(round(float(layout_values['purchase_sheet_ups'])))
                except ValueError:
                    pass
                planning_job.print_sheet_size = normalize_sheet_size(layout_values['print_sheet_size'])
                planning_job.purchase_sheet_size = normalize_sheet_size(layout_values['purchase_sheet_size'])
                planning_job.plate_set_no = layout_values['set_no'] or layout_values['new_set_no']
                if is_plate_ink_spec(planning_job.color_spec):
                    planning_job.color_spec = ''
                if resolved_print_color:
                    apply_print_color_to_planning_job(planning_job, resolved_print_color)
                planning_job.save()

            # Keep plate request AWC/set/die aligned with validated layout values.
            plate_request.awc_no = layout_values['awc_no']
            plate_request.set_no = layout_values['set_no']
            plate_request.new_set_no = layout_values['new_set_no']
            plate_request.die_cutting = normalize_die_cutting(layout_values.get('die_cutting'))

            if recipe and plate_request.sku_recipe_id != recipe.pk:
                plate_request.sku_recipe = recipe

            plate_request.status = PlateRequest.STATUS_SENT
            plate_request.sent_by = request.user
            plate_request.sent_at = timezone.now()
            plate_request.save()

            if recipe:
                recipe.refresh_from_db()
                missing_master = _missing_required_master_fields(recipe, recipe.job_name or (planning_job.job_name if planning_job else ''))
                if recipe.master_data_status == 'pending_review':
                    messages.success(
                        request,
                        f"Plate request sent to {vendor}. SKU master submitted for QC review.",
                    )
                elif missing_master:
                    messages.success(
                        request,
                        f"Plate request sent to {vendor}. Designer layout saved to SKU master (Draft). "
                        f"Planner must complete: {', '.join(missing_master)} before QC.",
                    )
                else:
                    messages.success(request, f"Plate request sent to {vendor}.")
            else:
                messages.success(request, f"Plate request sent to {vendor}.")
            
        elif action == 'receive_from_vendor':
            challan = request.POST.get('challan', '').strip()
            box = request.POST.get('box', '').strip()
            remarks = request.POST.get('remarks', '').strip()
            
            plate_request.challan = challan
            plate_request.box = box
            if remarks:
                plate_request.remarks = remarks
                
            plate_request.status = PlateRequest.STATUS_RECEIVED
            plate_request.received_by = request.user
            plate_request.received_at = timezone.now()
            plate_request.save()
            messages.success(request, "Plate request received from vendor.")
            
        elif action == 'issue_to_production':
            plate_request.status = PlateRequest.STATUS_AVAILABLE
            plate_request.save()  # Triggers save() hook to transition planning stage to plate_received
            if plate_request.is_replacement:
                messages.success(
                    request,
                    "Replacement plates issued to production. Job is Ready again (no longer Waiting for plate).",
                )
            else:
                messages.success(request, "Plates issued to production. Planning stage updated to Plate Received.")

        elif action == 'cancel_request':
            from django.core.exceptions import ValidationError
            from printing_plates.services import cancel_plate_request

            reason = (request.POST.get('cancel_reason') or request.POST.get('remarks') or '').strip()
            try:
                cancel_plate_request(plate_request, actor=request.user, reason=reason)
            except ValidationError as exc:
                message = exc.messages[0] if getattr(exc, 'messages', None) else str(exc)
                messages.error(request, message)
                return redirect('printing_plates:request_detail', pk=pk)
            messages.success(
                request,
                'Plate request cancelled (plates not required). '
                'Find it under Plate Requests → filter Cancelled (also on the planning job detail).',
            )
            return redirect(reverse('printing_plates:request_list') + '?type=cancelled')

        elif action == 'update_designer_layout':
            # Change management for designer fields while SKU master is not approved.
            from planning.services import (
                apply_designer_layout_to_sku_recipe,
                designer_layout_missing_fields,
                designer_layout_validation_errors,
                ensure_sku_recipe_for_planning_job,
                _missing_required_master_fields,
            )
            from core.print_colors import (
                apply_print_color_to_planning_job,
                apply_print_color_to_sku_recipe,
                resolve_print_color_name,
            )
            from printing_plates.constants import is_plate_ink_spec

            planning_job = plate_request.planning_job
            recipe = plate_request.sku_recipe
            if planning_job and not recipe:
                recipe = ensure_sku_recipe_for_planning_job(planning_job, actor=request.user)
                if recipe:
                    plate_request.sku_recipe = recipe
                    plate_request.save(update_fields=['sku_recipe'])

            if not recipe:
                messages.error(request, 'SKU master row was not found for this plate request.')
                return redirect('printing_plates:request_detail', pk=pk)
            if (recipe.master_data_status or '') == 'approved':
                messages.error(
                    request,
                    'SKU master is approved and locked. Ask planner/admin to Reopen SKU before correcting designer fields.',
                )
                return redirect('printing_plates:request_detail', pk=pk)

            layout_values = {
                'size_w_mm': request.POST.get('size_w_mm', '').strip(),
                'size_h_mm': request.POST.get('size_h_mm', '').strip(),
                'print_sheet_size': request.POST.get('print_sheet_size', '').strip(),
                'ups': request.POST.get('ups', '').strip(),
                'purchase_sheet_size': request.POST.get('purchase_sheet_size', '').strip(),
                'purchase_sheet_ups': request.POST.get('purchase_sheet_ups', '').strip(),
                'die_cutting': request.POST.get('die_cutting', '').strip(),
                'awc_no': normalize_awc_no(request.POST.get('awc_no', '')),
                'set_no': request.POST.get('set_no', plate_request.set_no or '').strip(),
                'new_set_no': request.POST.get('new_set_no', plate_request.new_set_no or '').strip(),
            }
            missing_designer = designer_layout_missing_fields(layout_values)
            layout_errors = designer_layout_validation_errors(layout_values)
            if missing_designer or layout_errors:
                messages.error(
                    request,
                    'Designer fields are required (except remarks): '
                    + ', '.join(missing_designer + layout_errors)
                    + '.',
                )
                return redirect('printing_plates:request_detail', pk=pk)

            print_color = request.POST.get('print_color', '').strip()
            resolved_print_color = resolve_print_color_name(print_color)
            if not resolved_print_color:
                for candidate in (planning_job.color_spec if planning_job else '', recipe.color_spec):
                    candidate = (candidate or '').strip()
                    if candidate and not is_plate_ink_spec(candidate):
                        resolved_print_color = resolve_print_color_name(candidate)
                        if resolved_print_color:
                            break
            if not resolved_print_color:
                messages.error(request, 'Print Color is required from the master list.')
                return redirect('printing_plates:request_detail', pk=pk)

            try:
                apply_designer_layout_to_sku_recipe(planning_job, recipe, layout_values)
            except ValueError as exc:
                messages.error(request, str(exc))
                return redirect('printing_plates:request_detail', pk=pk)

            plate_request.awc_no = layout_values['awc_no']
            plate_request.set_no = layout_values['set_no']
            plate_request.new_set_no = layout_values['new_set_no']
            plate_request.die_cutting = normalize_die_cutting(layout_values.get('die_cutting'))
            plate_request.save(update_fields=['awc_no', 'set_no', 'new_set_no', 'die_cutting', 'updated_at'])

            if planning_job:
                try:
                    planning_job.size_w_mm = int(round(float(layout_values['size_w_mm'])))
                    planning_job.size_h_mm = int(round(float(layout_values['size_h_mm'])))
                    planning_job.ups = int(round(float(layout_values['ups'])))
                    planning_job.purchase_sheet_ups = int(round(float(layout_values['purchase_sheet_ups'])))
                except ValueError:
                    pass
                planning_job.print_sheet_size = layout_values['print_sheet_size']
                planning_job.purchase_sheet_size = layout_values['purchase_sheet_size']
                planning_job.plate_set_no = layout_values['set_no'] or layout_values['new_set_no']
                if is_plate_ink_spec(planning_job.color_spec):
                    planning_job.color_spec = ''
                apply_print_color_to_planning_job(planning_job, resolved_print_color)
                planning_job.save()

            apply_print_color_to_sku_recipe(recipe, resolved_print_color)
            recipe.save(update_fields=['color_spec', 'updated_at'])
            recipe.refresh_from_db()

            missing_master = _missing_required_master_fields(
                recipe,
                recipe.job_name or (planning_job.job_name if planning_job else ''),
            )
            if recipe.master_data_status == 'pending_review':
                messages.success(request, 'Designer fields updated on SKU master and submitted for QC review.')
            elif missing_master:
                messages.success(
                    request,
                    'Designer fields updated on SKU master (Draft). Planner must complete: '
                    + ', '.join(missing_master)
                    + ' before QC.',
                )
            else:
                messages.success(request, 'Designer fields updated on SKU master.')
            return redirect('printing_plates:request_detail', pk=pk)

        return redirect('printing_plates:request_detail', pk=pk)


class PlateRequestCreateView(LoginRequiredMixin, GraphicsDesignerAccessMixin, CreateView):
    model = PlateRequest
    form_class = PlateRequestForm
    template_name = 'printing_plates/plate_request_form.html'
    success_url = reverse_lazy('printing_plates:queue')

    def get_form(self, form_class=None):
        form = super().get_form(form_class)
        form.fields['planning_job'].queryset = PlanningJob.objects.filter(
            planning_stage__in=['new_plate_making', 'repeat_plate_making']
        ).order_by('jc_number')
        return form

    def form_valid(self, form):
        form.instance.requested_by = self.request.user
        form.instance.requested_at = form.instance.requested_at or timezone.now()
        return super().form_valid(form)


class AdminAccessMixin(UserPassesTestMixin):
    def test_func(self):
        return _user_is_admin(self.request.user)


class BulkCancelStalePlateRequestsView(LoginRequiredMixin, AdminAccessMixin, View):
    def post(self, request):
        dry_run = (request.POST.get('dry_run') or '').strip() == '1'
        result = bulk_cancel_stale_open_plate_requests(actor=request.user, dry_run=dry_run)
        if dry_run:
            messages.info(
                request,
                f'Dry run: {result["total"]} stale open plate request(s) would be cancelled/archived.',
            )
        else:
            messages.success(
                request,
                f'Cancelled and archived {result["cancelled"]} of {result["total"]} stale open plate request(s). '
                f'Use Released Jobs for replacement plates on production jobs.',
            )
            if result.get('errors'):
                messages.warning(
                    request,
                    f'{len(result["errors"])} request(s) could not be cancelled. Check logs for details.',
                )
        next_url = (request.POST.get('next') or '').strip()
        if next_url:
            return redirect(next_url)
        return redirect(reverse('printing_plates:request_list') + '?type=cancelled')
