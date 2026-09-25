from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse

from .forms import (
    FlexoJobCardForm, FlexoPlanningDispatchRunFormSet, FlexoPlanningJobForm,
    FlexoPrintingLogEntryForm, FlexoSlittingCuttingLogEntryForm,
)
from .models import FLEXO_STATUS_CHOICES, FlexoJobCard, FlexoPlanningJob, MATERIAL_FAMILY_CHOICES
from .numbering import allocate_next_flexo_jc_number

FAMILY_LABELS = dict(MATERIAL_FAMILY_CHOICES)


def _family_or_404(family):
    if family not in FAMILY_LABELS:
        from django.http import Http404
        raise Http404("Unknown Flexo material family")
    return family


@login_required
def flexo_home(request):
    context = {
        'fl_count': FlexoPlanningJob.objects.filter(material_family='fl', is_active=True).count(),
        'satin_count': FlexoPlanningJob.objects.filter(material_family='satin', is_active=True).count(),
    }
    return render(request, 'flexo/flexo_home.html', context)


@login_required
def planning_list(request, family):
    _family_or_404(family)
    jobs = FlexoPlanningJob.objects.filter(material_family=family, is_active=True).select_related(
        'machine_name', 'department'
    )
    status = request.GET.get('status')
    if status:
        jobs = jobs.filter(status=status)
    q = request.GET.get('q')
    if q:
        jobs = jobs.filter(sku__icontains=q) | jobs.filter(jc_number__icontains=q) | jobs.filter(po_number__icontains=q)
    return render(request, 'flexo/planning_list.html', {
        'family': family,
        'family_label': FAMILY_LABELS[family],
        'jobs': jobs,
        'status_choices': FLEXO_STATUS_CHOICES,
    })


@login_required
def planning_create(request, family):
    _family_or_404(family)
    if request.method == 'POST':
        form = FlexoPlanningJobForm(request.POST)
        if form.is_valid():
            job = form.save(commit=False)
            job.material_family = family
            job.jc_number = allocate_next_flexo_jc_number(family)
            job.created_by = request.user
            job.save()
            messages.success(request, f'Planning job {job.jc_number} created.')
            return redirect('flexo:planning_detail', pk=job.pk)
    else:
        form = FlexoPlanningJobForm()
    return render(request, 'flexo/planning_form.html', {
        'family': family,
        'family_label': FAMILY_LABELS[family],
        'form': form,
        'is_new': True,
    })


@login_required
def planning_edit(request, pk):
    job = get_object_or_404(FlexoPlanningJob, pk=pk)
    if request.method == 'POST':
        form = FlexoPlanningJobForm(request.POST, instance=job)
        formset = FlexoPlanningDispatchRunFormSet(request.POST, instance=job)
        if form.is_valid() and formset.is_valid():
            form.save()
            formset.save()
            messages.success(request, f'Planning job {job.jc_number} updated.')
            return redirect('flexo:planning_detail', pk=job.pk)
    else:
        form = FlexoPlanningJobForm(instance=job)
        formset = FlexoPlanningDispatchRunFormSet(instance=job)
    return render(request, 'flexo/planning_form.html', {
        'family': job.material_family,
        'family_label': FAMILY_LABELS[job.material_family],
        'form': form,
        'formset': formset,
        'job': job,
        'is_new': False,
    })


@login_required
def planning_detail(request, pk):
    job = get_object_or_404(FlexoPlanningJob, pk=pk)
    return render(request, 'flexo/planning_detail.html', {
        'job': job,
        'family_label': FAMILY_LABELS[job.material_family],
        'dispatch_runs': job.dispatch_runs.all(),
    })


@login_required
def job_card_create(request, planning_pk):
    job = get_object_or_404(FlexoPlanningJob, pk=planning_pk)
    if hasattr(job, 'job_card'):
        return redirect('flexo:job_card_detail', pk=job.job_card.pk)

    card = FlexoJobCard.objects.create(
        planning_job=job,
        job_card_no=job.jc_number,
        material_family=job.material_family,
        sku=job.sku,
        job_name=job.job_name,
        po_number=job.po_number,
        po_received_date=job.po_received_date,
        order_qty=job.order_qty,
        delivery_location=job.destination,
        department=job.department,
        daily_demand=job.daily_demand,
        material_type=job.material,
        size_w_mm=job.size_w_mm,
        size_h_mm=job.size_h_mm,
        size_h_inches=job.size_h_inches,
        print_roll_size=job.roll_size,
        roll_meter_qty=job.roll_qty_meter,
        ups=job.ups,
        machine=job.machine_name,
        wastage=job.wastage_qty_meter,
        color=job.color_spec,
        mtr_per_roll=str(job.meter_per_roll) if job.meter_per_roll is not None else '',
        die_code=job.die_code,
        awc_no=job.awc_no,
        printing_cylinder=job.printing_cylinder,
    )
    messages.success(request, f'Job Card {card.job_card_no} created.')
    return redirect('flexo:job_card_detail', pk=card.pk)


@login_required
def job_card_detail(request, pk):
    card = get_object_or_404(FlexoJobCard, pk=pk)
    if request.method == 'POST':
        form = FlexoJobCardForm(request.POST, instance=card)
        if form.is_valid():
            form.save()
            messages.success(request, f'Job Card {card.job_card_no} updated.')
            return redirect('flexo:job_card_detail', pk=card.pk)
    else:
        form = FlexoJobCardForm(instance=card)
    return render(request, 'flexo/job_card_detail.html', {
        'card': card,
        'form': form,
        'printing_entries': card.printing_log_entries.all(),
        'slitting_entries': card.slitting_cutting_log_entries.all() if card.is_satin else None,
        'printing_form': FlexoPrintingLogEntryForm(),
        'slitting_form': FlexoSlittingCuttingLogEntryForm() if card.is_satin else None,
    })


@login_required
def job_card_add_printing_entry(request, pk):
    card = get_object_or_404(FlexoJobCard, pk=pk)
    if request.method == 'POST':
        form = FlexoPrintingLogEntryForm(request.POST)
        if form.is_valid():
            entry = form.save(commit=False)
            entry.job_card = card
            entry.save()
            messages.success(request, 'Printing log entry added.')
    return redirect('flexo:job_card_detail', pk=card.pk)


@login_required
def job_card_add_slitting_entry(request, pk):
    card = get_object_or_404(FlexoJobCard, pk=pk)
    if not card.is_satin:
        from django.http import Http404
        raise Http404("Slitting/Cutting log is only available for SATIN job cards")
    if request.method == 'POST':
        form = FlexoSlittingCuttingLogEntryForm(request.POST)
        if form.is_valid():
            entry = form.save(commit=False)
            entry.job_card = card
            entry.save()
            messages.success(request, 'Slitting/Cutting log entry added.')
    return redirect('flexo:job_card_detail', pk=card.pk)


@login_required
def job_card_print(request, pk):
    card = get_object_or_404(FlexoJobCard, pk=pk)
    return render(request, 'flexo/job_card_print.html', {
        'card': card,
        'printing_entries': card.printing_log_entries.all(),
        'slitting_entries': card.slitting_cutting_log_entries.all() if card.is_satin else None,
    })
