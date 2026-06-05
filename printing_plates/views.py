from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.urls import reverse_lazy
from django.utils import timezone
from django.views.generic import ListView, DetailView, CreateView, TemplateView

from planning.models import PlanningJob
from .models import PlateRequest


class GraphicsDesignerAccessMixin(UserPassesTestMixin):
    def test_func(self):
        profile = getattr(self.request.user, 'profile', None)
        return bool(profile and profile.can_view_plate_queue())


class PlateQueueView(LoginRequiredMixin, GraphicsDesignerAccessMixin, ListView):
    model = PlanningJob
    template_name = 'printing_plates/plate_queue.html'
    context_object_name = 'planning_jobs'

    def get_queryset(self):
        return PlanningJob.objects.filter(
            planning_stage__in=['new_plate_making', 'repeat_plate_making']
        ).select_related('job_card')


class PlateStatusBoardView(LoginRequiredMixin, GraphicsDesignerAccessMixin, ListView):
    model = PlateRequest
    template_name = 'printing_plates/plate_status_board.html'
    context_object_name = 'plate_requests'

    def get_queryset(self):
        return PlateRequest.objects.select_related(
            'planning_job', 'job_card', 'sku_recipe', 'machine', 'department'
        )


class PlateRequestDetailView(LoginRequiredMixin, GraphicsDesignerAccessMixin, DetailView):
    model = PlateRequest
    template_name = 'printing_plates/plate_request_detail.html'
    context_object_name = 'plate_request'


class PlateRequestCreateView(LoginRequiredMixin, GraphicsDesignerAccessMixin, CreateView):
    model = PlateRequest
    template_name = 'printing_plates/plate_request_form.html'
    fields = [
        'planning_job',
        'job_card',
        'sku_recipe',
        'machine',
        'department',
        'set_no',
        'new_set_no',
        'plate_quantity',
        'plate_color',
        'vendor',
        'remarks',
        'source',
        'challan',
        'chalan_sign',
        'box',
        'image',
        'link',
    ]
    success_url = reverse_lazy('printing_plates:queue')

    def form_valid(self, form):
        form.instance.requested_by = self.request.user
        form.instance.requested_at = form.instance.requested_at or timezone.now()
        return super().form_valid(form)
