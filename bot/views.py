from django.contrib import messages
from django.contrib.auth.decorators import login_required, user_passes_test
from django.core.paginator import Paginator
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone

from bot.forms import BotAutomationForm, TestSendForm
from bot.models import EXECUTION_STATUS_CHOICES, BotAutomation, BotExecution, BotGlobalSettings
from bot.services import (
    build_email_parts,
    refresh_next_run,
    resolve_recipients,
    run_bot_manually,
    send_test_email,
)
from bot.template_engine import SUPPORTED_VARIABLES


def can_manage_bots(user):
    """Same gate as the backup module (superuser or the admin role), plus the
    DB-driven nav.bot permission so access can be granted from
    Settings -> Roles & Permissions without code changes."""
    if not user.is_authenticated:
        return False
    if user.is_superuser:
        return True
    profile = getattr(user, 'profile', None)
    if profile and (getattr(profile, 'role', '') or '').strip().lower() == 'admin':
        return True
    from core.permissions import user_has_permission
    return user_has_permission(user, 'nav.bot')


bot_admin_required = user_passes_test(can_manage_bots)


@login_required
@bot_admin_required
def bot_list(request):
    bots = BotAutomation.objects.all().select_related('run_as')
    recent = BotExecution.objects.select_related('bot')[:10]
    return render(request, 'bot/bot_list.html', {
        'bots': bots,
        'recent_executions': recent,
        'active_count': sum(1 for bot in bots if bot.is_active),
        'global_settings': BotGlobalSettings.get_settings(),
    })


@login_required
@bot_admin_required
def bot_global_toggle(request):
    if request.method != 'POST':
        return redirect('bot:bot_list')

    settings_obj = BotGlobalSettings.get_settings()
    settings_obj.automation_enabled = not settings_obj.automation_enabled
    settings_obj.updated_by = request.user
    settings_obj.save(update_fields=['automation_enabled', 'updated_by', 'updated_at'])

    if settings_obj.automation_enabled:
        messages.success(request, 'Bot automations resumed — active bots will fire on their next scheduled time.')
    else:
        messages.warning(request, 'Bot automations paused. No scheduled email will send until this is turned back on.')
    return redirect('bot:bot_list')


@login_required
@bot_admin_required
def bot_create(request):
    return _bot_form(request, None)


@login_required
@bot_admin_required
def bot_edit(request, pk):
    return _bot_form(request, get_object_or_404(BotAutomation, pk=pk))


def _bot_form(request, bot):
    if request.method == 'POST':
        form = BotAutomationForm(request.POST, instance=bot)
        if form.is_valid():
            saved = form.save(commit=False)
            if bot is None:
                saved.created_by = request.user
            saved.save()
            refresh_next_run(saved)
            messages.success(request, f'Automation "{saved.name}" saved.')
            return redirect('bot:bot_edit', pk=saved.pk)
        messages.error(request, 'Please correct the errors below.')
    else:
        form = BotAutomationForm(instance=bot)

    return render(request, 'bot/bot_form.html', {
        'form': form,
        'bot': bot,
        'variables': SUPPORTED_VARIABLES,
        'is_create': bot is None,
    })


@login_required
@bot_admin_required
def bot_preview(request, pk):
    """Render the exact email in-browser without sending it."""
    bot = get_object_or_404(BotAutomation, pk=pk)
    to, cc, bcc = resolve_recipients(bot)

    parts = None
    error = None
    try:
        parts = build_email_parts(bot)
    except Exception as exc:  # noqa: BLE001 — surfaced on the page, not raised
        error = f'{type(exc).__name__}: {exc}'

    return render(request, 'bot/bot_preview.html', {
        'bot': bot,
        'parts': parts,
        'error': error,
        'to': to,
        'cc': cc,
        'bcc': bcc,
        'test_form': TestSendForm(initial={'email': request.user.email}),
        'would_skip': bool(parts and parts['record_count'] == 0 and not bot.send_when_empty),
    })


@login_required
@bot_admin_required
def bot_test_send(request, pk):
    bot = get_object_or_404(BotAutomation, pk=pk)
    if request.method != 'POST':
        return redirect('bot:bot_preview', pk=pk)

    form = TestSendForm(request.POST)
    if not form.is_valid():
        messages.error(request, 'Enter a valid email address for the test send.')
        return redirect('bot:bot_preview', pk=pk)

    address = form.cleaned_data['email']
    execution = send_test_email(bot, address, actor=request.user)
    if execution.is_failure:
        messages.error(request, f'Test send failed: {execution.error_message.splitlines()[0]}')
    else:
        messages.success(request, f'Test email sent to {address}.')
    return redirect('bot:execution_detail', pk=execution.pk)


@login_required
@bot_admin_required
def bot_run_now(request, pk):
    bot = get_object_or_404(BotAutomation, pk=pk)
    if request.method != 'POST':
        return redirect('bot:bot_list')

    execution = run_bot_manually(bot, actor=request.user)
    if execution.is_failure:
        messages.error(request, f'Run failed: {execution.error_message.splitlines()[0]}')
    elif execution.status == 'SKIPPED':
        messages.info(request, 'No pending records — nothing was sent (Send when empty is off).')
    else:
        messages.success(
            request,
            f'Sent to {execution.recipients_to or "(cc/bcc only)"} — {execution.record_count} record(s).',
        )
    return redirect('bot:execution_detail', pk=execution.pk)


@login_required
@bot_admin_required
def bot_toggle(request, pk):
    bot = get_object_or_404(BotAutomation, pk=pk)
    if request.method != 'POST':
        return redirect('bot:bot_list')

    if not bot.is_active:
        to, cc, bcc = resolve_recipients(bot)
        if not (to or cc or bcc):
            messages.error(request, 'Add at least one recipient before activating this automation.')
            return redirect('bot:bot_edit', pk=pk)

    bot.is_active = not bot.is_active
    bot.save(update_fields=['is_active', 'updated_at'])
    next_run = refresh_next_run(bot)

    if bot.is_active:
        when = timezone.localtime(next_run).strftime('%d %b %Y, %I:%M %p') if next_run else 'never (check dates)'
        messages.success(request, f'"{bot.name}" activated. Next run: {when}.')
    else:
        messages.info(request, f'"{bot.name}" deactivated.')
    return redirect('bot:bot_list')


@login_required
@bot_admin_required
def execution_list(request):
    executions = BotExecution.objects.select_related('bot', 'triggered_by')

    bot_id = (request.GET.get('bot') or '').strip()
    status = (request.GET.get('status') or '').strip()
    if bot_id.isdigit():
        executions = executions.filter(bot_id=int(bot_id))
    if status:
        executions = executions.filter(status=status)

    page = Paginator(executions, 50).get_page(request.GET.get('page'))
    return render(request, 'bot/execution_list.html', {
        'page_obj': page,
        'bots': BotAutomation.objects.all(),
        'status_choices': EXECUTION_STATUS_CHOICES,
        'selected_bot': bot_id,
        'selected_status': status,
    })


@login_required
@bot_admin_required
def execution_detail(request, pk):
    execution = get_object_or_404(
        BotExecution.objects.select_related('bot', 'triggered_by'), pk=pk
    )
    return render(request, 'bot/execution_detail.html', {'execution': execution})
