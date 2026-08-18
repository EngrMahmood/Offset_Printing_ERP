from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.views.decorators.http import require_POST
from django.contrib.auth.models import User
from django.core.paginator import Paginator
from django.db.models import Avg, Count, Q
from django.utils import timezone
from .models import REMIND_FROM_CHOICES, Team, Task, TaskComment, TaskNotificationLog, TaskNotificationSettings
from .forms import TaskForm, TeamForm

def is_manager_or_admin(user):
    profile = getattr(user, 'profile', None)
    if not profile:
        return user.is_supertype or user.is_staff
    from core.permissions import user_has_permission
    return user_has_permission(user, 'action.manage_tasks')

@login_required
def dashboard(request):
    # Filter variables
    status_filter = request.GET.get('status', '')
    priority_filter = request.GET.get('priority', '')
    assignee_filter = request.GET.get('assignee', '')
    team_filter = request.GET.get('team', '')
    
    # Query tasks
    tasks = Task.objects.all().select_related('assignee', 'assigned_team', 'created_by')
    
    if status_filter:
        tasks = tasks.filter(status=status_filter)
    if priority_filter:
        tasks = tasks.filter(priority=priority_filter)
    if assignee_filter:
        tasks = tasks.filter(assignee_id=assignee_filter)
    if team_filter:
        tasks = tasks.filter(assigned_team_id=team_filter)
        
    # Leaderboards
    # Individual Leaderboard
    employee_leaderboard = User.objects.annotate(
        completed_count=Count('assigned_tasks', filter=Q(assigned_tasks__status__in=['completed', 'verified'])),
        avg_score=Avg('assigned_tasks__score', filter=Q(assigned_tasks__status__in=['completed', 'verified']))
    ).filter(completed_count__gt=0).order_by('-avg_score', '-completed_count')
    
    # Team Leaderboard
    team_leaderboard = Team.objects.annotate(
        completed_count=Count('assigned_tasks', filter=Q(assigned_tasks__status__in=['completed', 'verified'])),
        avg_score=Avg('assigned_tasks__score', filter=Q(assigned_tasks__status__in=['completed', 'verified']))
    ).filter(completed_count__gt=0).order_by('-avg_score', '-completed_count')

    # Fetch users and teams for filtering dropdowns
    users = User.objects.filter(is_active=True).order_by('username')
    teams = Team.objects.all()

    context = {
        'tasks': tasks[:200],
        'employee_leaderboard': employee_leaderboard[:10],
        'team_leaderboard': team_leaderboard[:10],
        'users': users,
        'teams': teams,
        'status_filter': status_filter,
        'priority_filter': priority_filter,
        'assignee_filter': assignee_filter,
        'team_filter': team_filter,
        'is_manager': is_manager_or_admin(request.user),
    }
    return render(request, 'tasks/dashboard.html', context)


@login_required
def create_task(request):
    if request.method == 'POST':
        form = TaskForm(request.POST)
        if form.is_valid():
            task = form.save(commit=False)
            task.created_by = request.user
            task.save()
            messages.success(request, f"Task '{task.title}' created and assigned successfully!")
            return redirect('tasks:dashboard')
    else:
        form = TaskForm()
        
    return render(request, 'tasks/form.html', {'form': form, 'title': 'Create Task'})


@login_required
def edit_task(request, pk):
    task = get_object_or_404(Task, pk=pk)
    
    # Only creator, managers, or admins can edit tasks
    if task.created_by != request.user and not is_manager_or_admin(request.user):
        messages.error(request, "You do not have permission to edit this task.")
        return redirect('tasks:detail', pk=pk)
        
    if request.method == 'POST':
        form = TaskForm(request.POST, instance=task)
        if form.is_valid():
            task = form.save()
            messages.success(request, f"Task '{task.title}' updated successfully!")
            return redirect('tasks:detail', pk=pk)
    else:
        form = TaskForm(instance=task)
        
    return render(request, 'tasks/form.html', {'form': form, 'title': 'Edit Task', 'task': task})


@login_required
def delete_task(request, pk):
    task = get_object_or_404(Task, pk=pk)
    
    # Only creator, managers, or admins can delete tasks
    if task.created_by != request.user and not is_manager_or_admin(request.user):
        messages.error(request, "You do not have permission to delete this task.")
        return redirect('tasks:detail', pk=pk)
        
    if request.method == 'POST':
        task.delete()
        messages.success(request, "Task deleted successfully.")
        return redirect('tasks:dashboard')
        
    return render(request, 'confirm_delete.html', {'entity_type': 'Task', 'record_label': task.title})


@login_required
def task_detail(request, pk):
    task = get_object_or_404(Task, pk=pk)
    comments = task.comments.all().select_related('user')
    
    # Check permissions
    can_edit = task.created_by == request.user or is_manager_or_admin(request.user)
    is_assignee = task.assignee == request.user or (task.assigned_team and request.user in task.assigned_team.members.all())
    can_grade = is_manager_or_admin(request.user) or task.created_by == request.user

    if request.method == 'POST':
        # Add Comment
        comment_text = request.POST.get('comment', '').strip()
        if comment_text:
            TaskComment.objects.create(task=task, user=request.user, comment=comment_text)
            messages.success(request, "Comment added.")
            return redirect('tasks:detail', pk=pk)

    context = {
        'task': task,
        'comments': comments,
        'can_edit': can_edit,
        'is_assignee': is_assignee,
        'can_grade': can_grade,
    }
    return render(request, 'tasks/detail.html', context)


@login_required
@require_POST
def update_status(request, pk):
    task = get_object_or_404(Task, pk=pk)
    new_status = request.POST.get('status', '').strip()
    
    is_assignee = task.assignee == request.user or (task.assigned_team and request.user in task.assigned_team.members.all())
    is_creator_or_manager = task.created_by == request.user or is_manager_or_admin(request.user)
    
    if not is_assignee and not is_creator_or_manager:
        messages.error(request, "You are not authorized to update the status of this task.")
        return redirect('tasks:detail', pk=pk)
        
    if new_status in dict(Task.STATUS_CHOICES):
        # Enforce that only managers/creators can mark task as 'verified'
        if new_status == 'verified' and not is_creator_or_manager:
            messages.error(request, "Only the task creator or a manager can verify completion.")
            return redirect('tasks:detail', pk=pk)
            
        task.status = new_status
        task.save()
        messages.success(request, f"Task status updated to '{task.get_status_display()}'.")
        
    return redirect('tasks:detail', pk=pk)


@login_required
@require_POST
def grade_task(request, pk):
    task = get_object_or_404(Task, pk=pk)
    
    # Only creator, managers, or admins can grade/score tasks
    if task.created_by != request.user and not is_manager_or_admin(request.user):
        messages.error(request, "You are not authorized to grade this task.")
        return redirect('tasks:detail', pk=pk)
        
    score_raw = request.POST.get('score', '').strip()
    score_remarks = request.POST.get('score_remarks', '').strip()
    
    try:
        score = int(score_raw)
        if not (0 <= score <= 100):
            raise ValueError()
    except ValueError:
        messages.error(request, "Score must be a number between 0 and 100.")
        return redirect('tasks:detail', pk=pk)
        
    task.score = score
    task.score_remarks = score_remarks
    task.scored_by = request.user
    # Force verification status when graded
    if task.status != 'verified':
        task.status = 'verified'
    task.save()
    
    messages.success(request, f"Task graded successfully with a score of {score}!")
    return redirect('tasks:detail', pk=pk)


@login_required
def teams_list(request):
    teams = Team.objects.all().prefetch_related('members')
    
    if request.method == 'POST':
        # Create team
        if not is_manager_or_admin(request.user):
            messages.error(request, "Only managers or admins can create teams.")
            return redirect('tasks:teams')
            
        form = TeamForm(request.POST)
        if form.is_valid():
            team = form.save()
            messages.success(request, f"Team '{team.name}' created successfully!")
            return redirect('tasks:teams')
    else:
        form = TeamForm()
        
    context = {
        'teams': teams,
        'form': form,
        'is_manager': is_manager_or_admin(request.user),
    }
    return render(request, 'tasks/teams.html', context)


@login_required
def automation(request):
    """Global assignment/reminder-email defaults + the notification activity
    log — manager/admin only, same gate as edit/delete/grade elsewhere in
    this app (not superuser-only, unlike the old site-wide Settings page)."""
    if not is_manager_or_admin(request.user):
        messages.error(request, "You do not have permission to manage task automation.")
        return redirect('tasks:dashboard')

    settings_obj = TaskNotificationSettings.get_solo()

    if request.method == 'POST':
        interval_raw = (request.POST.get('reminder_interval_days') or '').strip()
        try:
            interval_days = int(interval_raw)
            if interval_days < 1:
                raise ValueError()
        except ValueError:
            messages.error(request, 'Reminder interval must be a whole number of days, at least 1.')
            return redirect('tasks:automation')

        remind_from = (request.POST.get('remind_from') or '').strip()
        if remind_from not in dict(REMIND_FROM_CHOICES):
            messages.error(request, 'Invalid "start reminders from" value.')
            return redirect('tasks:automation')

        settings_obj.assignment_email_enabled = bool(request.POST.get('assignment_email_enabled'))
        settings_obj.reminders_enabled = bool(request.POST.get('reminders_enabled'))
        settings_obj.reminder_interval_days = interval_days
        settings_obj.remind_from = remind_from
        settings_obj.updated_by = request.user
        settings_obj.save()

        messages.success(request, 'Task automation settings saved.')
        return redirect('tasks:automation')

    logs = TaskNotificationLog.objects.select_related('task').all()
    paginator = Paginator(logs, 50)
    page_obj = paginator.get_page(request.GET.get('page'))

    context = {
        'settings': settings_obj,
        'page_obj': page_obj,
        'is_manager': True,
    }
    return render(request, 'tasks/automation.html', context)
