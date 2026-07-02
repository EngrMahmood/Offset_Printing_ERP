from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.db.models import Prefetch, Q
from django.urls import reverse_lazy
from django.utils import timezone
from django.contrib import messages
from django.views.generic import ListView, DetailView, CreateView, TemplateView

from planning.models import PlanningJob
from .forms import PlateRequestForm
from .models import PlateRequest
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
        req_qs = PlateRequest.objects.all()
        context['req_all_count'] = req_qs.count()
        context['req_repeat_count'] = req_qs.filter(
            Q(planning_job__repeat_flag='Repeat') | Q(job_card__planning_job__repeat_flag='Repeat')
        ).count()
        context['req_new_artwork_count'] = req_qs.filter(
            Q(planning_job__repeat_flag='New') | Q(job_card__planning_job__repeat_flag='New')
        ).count()
        
        # 2. Plate Sent Counts
        sent_qs = PlateRequest.objects.filter(status=PlateRequest.STATUS_SENT)
        context['sent_all_count'] = sent_qs.count()
        
        vendors = ['Dot Max', 'Ali Print Pack', 'Daniyal Process', 'In House', 'Cancel']
        sent_counts = {}
        for v in vendors:
            sent_counts[v.replace(' ', '_')] = sent_qs.filter(vendor=v).count()
        context['sent_counts'] = sent_counts
        
        # 3. Plate Received Counts
        rec_qs = PlateRequest.objects.filter(status=PlateRequest.STATUS_RECEIVED)
        context['rec_all_count'] = rec_qs.count()
        context['rec_empty_count'] = rec_qs.filter(vendor='').count()
        
        rec_counts = {}
        for v in vendors:
            rec_counts[v.replace(' ', '_')] = rec_qs.filter(vendor=v).count()
        context['rec_counts'] = rec_counts
        
        return context


class PlateRequestListView(LoginRequiredMixin, GraphicsDesignerAccessMixin, ListView):
    model = PlateRequest
    template_name = 'printing_plates/plate_request_list.html'
    context_object_name = 'plate_requests'
    paginate_by = 50

    def get_queryset(self):
        queryset = PlateRequest.objects.select_related(
            'planning_job', 'job_card', 'sku_recipe', 'machine', 'department', 'requested_by'
        ).order_by('-requested_at', '-created_at')

        # Type filter from AppSheet sidebar
        self.request_type = self.request.GET.get('type', '').strip().lower()
        if self.request_type == 'repeat':
            queryset = queryset.filter(
                Q(planning_job__repeat_flag='Repeat') | Q(job_card__planning_job__repeat_flag='Repeat')
            )
        elif self.request_type == 'new_artwork':
            queryset = queryset.filter(
                Q(planning_job__repeat_flag='New') | Q(job_card__planning_job__repeat_flag='New')
            )
        elif self.request_type == 'empty':
            queryset = queryset.exclude(
                Q(planning_job__repeat_flag='New') | Q(planning_job__repeat_flag='Repeat') |
                Q(job_card__planning_job__repeat_flag='New') | Q(job_card__planning_job__repeat_flag='Repeat')
            )

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
        
        # Calculate dynamic counts for sidebar
        base_qs = PlateRequest.objects.all()
        
        context['repeat_count'] = base_qs.filter(
            Q(planning_job__repeat_flag='Repeat') | Q(job_card__planning_job__repeat_flag='Repeat')
        ).count()
        
        context['new_artwork_count'] = base_qs.filter(
            Q(planning_job__repeat_flag='New') | Q(job_card__planning_job__repeat_flag='New')
        ).count()
        
        context['empty_count'] = base_qs.exclude(
            Q(planning_job__repeat_flag='New') | Q(planning_job__repeat_flag='Repeat') |
            Q(job_card__planning_job__repeat_flag='New') | Q(job_card__planning_job__repeat_flag='Repeat')
        ).count()

        context['all_count'] = base_qs.count()
        context['request_type'] = self.request_type
        context['q'] = self.request.GET.get('q', '')
        context['list_count'] = self.get_queryset().count()
        return context


class PlateQueueView(LoginRequiredMixin, GraphicsDesignerAccessMixin, ListView):
    model = PlanningJob
    template_name = 'printing_plates/plate_queue.html'
    context_object_name = 'planning_jobs'
    paginate_by = 50

    def get_queryset(self):
        active_statuses = [
            PlateRequest.STATUS_DRAFT,
            PlateRequest.STATUS_SENT,
            PlateRequest.STATUS_RECEIVED,
        ]
        active_plate_requests = Prefetch(
            'plate_requests',
            queryset=PlateRequest.objects.order_by('-requested_at', '-created_at'),
        )
        queryset = PlanningJob.objects.filter(
            planning_stage__in=['new_plate_making', 'repeat_plate_making']
        ).filter(
            Q(plate_requests__status__in=active_statuses) | Q(plate_requests__isnull=True)
        ).select_related('job_card').prefetch_related(active_plate_requests).distinct().order_by('-updated_at', '-id')

        # Search filter
        q = self.request.GET.get('q', '').strip()
        if q:
            queryset = queryset.filter(
                Q(jc_number__icontains=q) |
                Q(job_name__icontains=q) |
                Q(sku__icontains=q)
            )

        # Stage filter
        stage = self.request.GET.get('stage', '').strip()
        if stage in ['new_plate_making', 'repeat_plate_making']:
            queryset = queryset.filter(planning_stage=stage)

        return queryset

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['q'] = self.request.GET.get('q', '')
        context['stage'] = self.request.GET.get('stage', '')
        page_obj = context.get('page_obj')
        if page_obj is not None:
            context['queue_count'] = page_obj.paginator.count
        else:
            context['queue_count'] = len(context.get('planning_jobs') or [])
        return context





class PlateSentListView(LoginRequiredMixin, GraphicsDesignerAccessMixin, ListView):
    model = PlateRequest
    template_name = 'printing_plates/plate_sent_list.html'
    context_object_name = 'plate_requests'
    paginate_by = 50

    def get_queryset(self):
        queryset = PlateRequest.objects.filter(status=PlateRequest.STATUS_SENT).select_related(
            'planning_job', 'job_card', 'sku_recipe', 'machine', 'department', 'requested_by', 'sent_by'
        ).order_by('-sent_at', '-requested_at')

        # Vendor filter from sidebar
        self.vendor_filter = self.request.GET.get('vendor', '').strip()
        if self.vendor_filter:
            queryset = queryset.filter(vendor=self.vendor_filter)

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
        
        # Base query of sent requests
        base_qs = PlateRequest.objects.filter(status=PlateRequest.STATUS_SENT)
        
        # Setup vendor counts
        vendors = ['Dot Max', 'Ali Print Pack', 'Daniyal Process', 'In House', 'Cancel']
        counts = {}
        for v in vendors:
            key_name = v.replace(' ', '_')
            counts[key_name] = base_qs.filter(vendor=v).count()
            
        context['vendor_counts'] = counts
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
        queryset = PlateRequest.objects.filter(status=PlateRequest.STATUS_RECEIVED).select_related(
            'planning_job', 'job_card', 'sku_recipe', 'machine', 'department', 'requested_by', 'received_by'
        ).order_by('-received_at', '-requested_at')

        # Vendor filter from sidebar
        self.vendor_filter = self.request.GET.get('vendor', '').strip()
        if self.vendor_filter == 'empty':
            queryset = queryset.filter(vendor='')
        elif self.vendor_filter:
            queryset = queryset.filter(vendor=self.vendor_filter)

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
        
        # Base query of received requests
        base_qs = PlateRequest.objects.filter(status=PlateRequest.STATUS_RECEIVED)
        
        # Setup vendor counts
        vendors = ['Dot Max', 'Ali Print Pack', 'Daniyal Process', 'In House', 'Cancel']
        counts = {}
        for v in vendors:
            key_name = v.replace(' ', '_')
            counts[key_name] = base_qs.filter(vendor=v).count()
            
        context['vendor_counts'] = counts
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

        master_recipe = plate_request.sku_recipe  # default: linked recipe
        if sku:
            from planning.models import SkuRecipe
            _STATUS_ORDER = {'approved': 0, 'reviewed': 1, 'pending_review': 2, 'draft': 3}
            all_recipes = list(SkuRecipe.objects.filter(sku__iexact=sku))
            if all_recipes:
                master_recipe = min(
                    all_recipes,
                    key=lambda r: _STATUS_ORDER.get(r.master_data_status or '', 99)
                )

        context['master_recipe'] = master_recipe
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
            awc_no = request.POST.get('awc_no', '').strip()
            plate_color = request.POST.get('plate_color', '').strip()
            plate_quantity = request.POST.get('plate_quantity', '').strip()
            remarks = request.POST.get('remarks', '').strip()
            
            if not vendor:
                messages.error(request, "Vendor is required.")
                return redirect('printing_plates:request_detail', pk=pk)
                
            plate_request.vendor = vendor
            plate_request.set_no = set_no
            plate_request.new_set_no = new_set_no
            plate_request.awc_no = awc_no
            plate_request.plate_color = plate_color
            if plate_quantity:
                try:
                    plate_request.plate_quantity = int(plate_quantity)
                except ValueError:
                    messages.error(request, "Plate Quantity must be a valid integer.")
                    return redirect('printing_plates:request_detail', pk=pk)
            plate_request.remarks = remarks

            planning_job = plate_request.planning_job
            if planning_job:
                planning_job.remarks = remarks
                planning_job.save(update_fields=['remarks'])

            # If sku_recipe was not linked at plate request creation time,
            # resolve the best available recipe now and link it.
            if not plate_request.sku_recipe:
                sku = (
                    (planning_job.sku if planning_job else None) or ''
                ).strip()
                if sku:
                    from planning.models import SkuRecipe as _SR
                    _STATUS_ORDER = {'approved': 0, 'reviewed': 1, 'pending_review': 2, 'draft': 3}
                    _candidates = list(_SR.objects.filter(sku__iexact=sku))
                    if _candidates:
                        plate_request.sku_recipe = min(
                            _candidates,
                            key=lambda r: _STATUS_ORDER.get(r.master_data_status or '', 99)
                        )
                        plate_request.save(update_fields=['sku_recipe'])

            recipe = plate_request.sku_recipe
            if recipe:
                recipe.remarks = remarks
                recipe.notes = remarks
                recipe.save(update_fields=['remarks', 'notes'])
                
            # Process Designer Layout specifications for New jobs
            if plate_request.planning_job and plate_request.planning_job.repeat_flag == 'New':
                size_w_mm = request.POST.get('size_w_mm', '').strip()
                size_h_mm = request.POST.get('size_h_mm', '').strip()
                print_sheet_size = request.POST.get('print_sheet_size', '').strip()
                ups = request.POST.get('ups', '').strip()
                purchase_sheet_size = request.POST.get('purchase_sheet_size', '').strip()
                purchase_sheet_ups = request.POST.get('purchase_sheet_ups', '').strip()
                die_cutting = request.POST.get('die_cutting', '').strip()

                if not (size_w_mm and size_h_mm and print_sheet_size and ups and purchase_sheet_size and purchase_sheet_ups and die_cutting):
                    messages.error(request, "All layout specification fields are required for new jobs.")
                    return redirect('printing_plates:request_detail', pk=pk)

                recipe = plate_request.sku_recipe
                if recipe:
                    try:
                        recipe.size_w_mm = int(round(float(size_w_mm)))
                        recipe.size_h_mm = int(round(float(size_h_mm)))
                    except ValueError:
                        messages.error(request, "Size Width and Height must be valid numbers.")
                        return redirect('printing_plates:request_detail', pk=pk)
                    
                    recipe.print_sheet_size = print_sheet_size
                    try:
                        recipe.ups = int(round(float(ups)))
                    except ValueError:
                        messages.error(request, "Ups must be a valid number.")
                        return redirect('printing_plates:request_detail', pk=pk)

                    recipe.purchase_sheet_size = purchase_sheet_size
                    try:
                        recipe.purchase_sheet_ups = int(round(float(purchase_sheet_ups)))
                    except ValueError:
                        messages.error(request, "Purchase Sheet Ups must be a valid number.")
                        return redirect('printing_plates:request_detail', pk=pk)

                    recipe.color_spec = plate_color
                    recipe.awc_no = awc_no
                    recipe.die_cutting = die_cutting
                    recipe.plate_set_no = set_no or new_set_no
                    recipe.master_data_status = 'pending_review'
                    recipe.save()

                planning_job = plate_request.planning_job
                if planning_job:
                    try:
                        planning_job.size_w_mm = int(round(float(size_w_mm)))
                        planning_job.size_h_mm = int(round(float(size_h_mm)))
                        planning_job.ups = int(round(float(ups)))
                        planning_job.purchase_sheet_ups = int(round(float(purchase_sheet_ups)))
                    except ValueError:
                        pass
                    
                    planning_job.print_sheet_size = print_sheet_size
                    planning_job.purchase_sheet_size = purchase_sheet_size
                    planning_job.color_spec = plate_color
                    planning_job.plate_set_no = set_no or new_set_no
                    planning_job.save()
            else:
                planning_job = plate_request.planning_job
                if planning_job:
                    planning_job.plate_set_no = set_no or new_set_no
                    planning_job.save()

            plate_request.status = PlateRequest.STATUS_SENT
            plate_request.sent_by = request.user
            plate_request.sent_at = timezone.now()
            plate_request.save()
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
            messages.success(request, "Plates issued to production. Planning stage updated to Plate Received.")
            
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
