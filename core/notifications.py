"""In-app notification engine (navbar bell + side toasts) with Rule-Based Routing."""

from __future__ import annotations

import logging
from django.contrib.auth import get_user_model
from django.urls import NoReverseMatch, reverse
from django.template import Template, Context
from django.utils import timezone

from core.models import Notification

User = get_user_model()
logger = logging.getLogger(__name__)


def users_with_roles(*roles):
    role_set = {str(role or '').strip().lower() for role in roles if role}
    if not role_set:
        return User.objects.none()
    return User.objects.filter(
        is_active=True,
        profile__role__in=role_set,
    ).distinct()


def _safe_reverse(name, *args, **kwargs):
    try:
        return reverse(name, args=args, kwargs=kwargs)
    except NoReverseMatch:
        return ''


def render_template_string(template_str, instance, actor, extra_context=None):
    if not template_str:
        return ''
    ctx = {
        'instance': instance,
        'actor': actor,
    }
    if extra_context:
        ctx.update(extra_context)
    try:
        t = Template(template_str)
        return t.render(Context(ctx)).strip()
    except Exception as e:
        logger.warning("Error rendering notification template: %s", e)
        return template_str


def resolve_creator(instance):
    if not instance:
        return None
    for attr in ['created_by', 'requested_by', 'user']:
        val = getattr(instance, attr, None)
        if val and getattr(val, 'is_active', False):
            return val
    created_by_id = getattr(instance, 'created_by_id', None)
    if created_by_id:
        try:
            u = User.objects.get(pk=created_by_id)
            if u.is_active:
                return u
        except Exception:
            pass
    return None


def resolve_next_stage_roles(instance):
    if not instance:
        return []

    module_name = instance.__class__.__name__
    if module_name == 'SkuRecipe':
        module_key = 'SKU'
        current_stage = getattr(instance, 'master_data_status', '')
    elif module_name == 'PlateRequest':
        module_key = 'Plate'
        current_stage = getattr(instance, 'status', '')
    elif module_name == 'JobCard':
        module_key = 'JobCard'
        current_stage = getattr(instance, 'workflow_status', '')
    else:
        module_key = module_name
        current_stage = ''

    if current_stage:
        try:
            from core.models import WorkflowTransition
            transitions = WorkflowTransition.objects.filter(
                module__iexact=module_key,
                current_stage__iexact=current_stage,
            )
            if transitions.exists():
                roles = [t.notify_role for t in transitions if t.notify_role]
                if roles:
                    return roles
        except Exception as e:
            logger.warning("Error querying WorkflowTransition: %s", e)

    if hasattr(instance, 'get_next_notification_roles'):
        try:
            return instance.get_next_notification_roles()
        except Exception:
            pass
    return []


def notify_event(event_code, instance=None, actor=None, extra_context=None):
    """
    Generic service to send notifications using database configuration.
    """
    event_code = (event_code or '').strip()
    if not event_code:
        return 0

    explicit_users = (extra_context or {}).get('explicit_users')
    explicit_roles = (extra_context or {}).get('explicit_roles')
    actor_id = getattr(actor, 'pk', actor) if actor else None

    # 1. Resolve recipients
    recipient_ids = set()

    if explicit_users is not None or explicit_roles is not None:
        # Backward compatibility/explicit routing
        resolved_users = set()
        if explicit_users:
            for u in explicit_users:
                if u is None:
                    continue
                resolved_users.add(u)
        if explicit_roles:
            for u in users_with_roles(*explicit_roles):
                resolved_users.add(u)

        exclude_act = (extra_context or {}).get('exclude_actor', True)
        for u in resolved_users:
            u_id = getattr(u, 'pk', None) or u
            if not u_id:
                continue
            if exclude_act and actor_id and int(u_id) == int(actor_id):
                continue
            recipient_ids.add(int(u_id))
    else:
        # Rule-based routing
        try:
            from core.models import NotificationEvent, NotificationRule
            event = NotificationEvent.objects.filter(code=event_code, is_active=True).first()
        except Exception as e:
            logger.exception("Error fetching NotificationEvent: %s", e)
            return 0

        if not event:
            logger.warning("NotificationEvent %s does not exist or is inactive.", event_code)
            return 0

        rules = NotificationRule.objects.filter(event=event, enabled=True)
        for rule in rules:
            rule_recipients = set()
            
            # Recipient type resolution
            if rule.recipient_type == 'role' and rule.role:
                for u in users_with_roles(rule.role):
                    rule_recipients.add(u)
            elif rule.recipient_type == 'user' and rule.user:
                if rule.user.is_active:
                    rule_recipients.add(rule.user)
            elif rule.recipient_type == 'department' and rule.department:
                from core.models import UserProfile
                profiles = UserProfile.objects.filter(department=rule.department, user__is_active=True).select_related('user')
                for p in profiles:
                    rule_recipients.add(p.user)
            elif rule.recipient_type == 'creator' or rule.send_to_creator:
                creator = resolve_creator(instance)
                if creator:
                    rule_recipients.add(creator)
            elif rule.recipient_type == 'manager' or rule.send_to_manager:
                creator = resolve_creator(instance)
                manager = getattr(creator.profile, 'manager', None) if creator and hasattr(creator, 'profile') else None
                if manager and manager.is_active:
                    rule_recipients.add(manager)
                else:
                    for u in users_with_roles('manager'):
                        rule_recipients.add(u)
            elif rule.recipient_type == 'supervisor' or rule.send_to_supervisor:
                creator = resolve_creator(instance)
                supervisor = getattr(creator.profile, 'supervisor', None) if creator and hasattr(creator, 'profile') else None
                if supervisor and supervisor.is_active:
                    rule_recipients.add(supervisor)
                else:
                    for u in users_with_roles('production', 'production_manager'):
                        rule_recipients.add(u)
            elif rule.recipient_type == 'next_stage' or rule.send_to_next_stage:
                roles = resolve_next_stage_roles(instance)
                if roles:
                    for u in users_with_roles(*roles):
                        rule_recipients.add(u)

            # Check boolean flags independently
            if rule.send_to_creator:
                creator = resolve_creator(instance)
                if creator:
                    rule_recipients.add(creator)
            if rule.send_to_manager:
                creator = resolve_creator(instance)
                manager = getattr(creator.profile, 'manager', None) if creator and hasattr(creator, 'profile') else None
                if manager and manager.is_active:
                    rule_recipients.add(manager)
                else:
                    for u in users_with_roles('manager'):
                        rule_recipients.add(u)
            if rule.send_to_supervisor:
                creator = resolve_creator(instance)
                supervisor = getattr(creator.profile, 'supervisor', None) if creator and hasattr(creator, 'profile') else None
                if supervisor and supervisor.is_active:
                    rule_recipients.add(supervisor)
                else:
                    for u in users_with_roles('production', 'production_manager'):
                        rule_recipients.add(u)
            if rule.send_to_next_stage:
                roles = resolve_next_stage_roles(instance)
                if roles:
                    for u in users_with_roles(*roles):
                        rule_recipients.add(u)

            # Apply actor exclusion
            for u in rule_recipients:
                u_id = getattr(u, 'pk', None) or u
                if not u_id:
                    continue
                if rule.exclude_actor and actor_id and int(u_id) == int(actor_id):
                    continue
                recipient_ids.add(int(u_id))

    if not recipient_ids:
        return 0

    # 2. Render templates
    event = None
    try:
        from core.models import NotificationEvent
        event = NotificationEvent.objects.filter(code=event_code, is_active=True).first()
    except Exception:
        pass

    title = (extra_context or {}).get('title') or ''
    message = (extra_context or {}).get('message') or ''
    link = (extra_context or {}).get('link') or ''
    entity_type = (extra_context or {}).get('entity_type') or ''
    entity_id = (extra_context or {}).get('entity_id')

    if event:
        if not title:
            title = render_template_string(event.title_template, instance, actor, extra_context)
        if not message:
            message = render_template_string(event.message_template, instance, actor, extra_context)
        if not link:
            link = render_template_string(event.link_template, instance, actor, extra_context)

    if not title:
        title = event_code

    if not entity_type and instance:
        entity_type = instance.__class__.__name__.lower()
    if not entity_id and instance:
        entity_id = instance.pk

    rows = [
        Notification(
            user_id=user_id,
            event_type=event_code[:80],
            title=title[:200],
            message=(message or '').strip(),
            link=(link or '').strip()[:500],
            entity_type=(entity_type or '').strip()[:60],
            entity_id=entity_id,
            created_by=actor if getattr(actor, 'pk', None) else None,
        )
        for user_id in recipient_ids
    ]
    Notification.objects.bulk_create(rows)
    return len(rows)


def notify_users(
    users,
    *,
    event_type,
    title,
    message='',
    link='',
    entity_type='',
    entity_id=None,
    actor=None,
    exclude_actor=True,
):
    """Wrapper calling generic notify_event engine for backward compatibility."""
    return notify_event(
        event_code=event_type,
        instance=None,
        actor=actor,
        extra_context={
            'explicit_users': users,
            'title': title,
            'message': message,
            'link': link,
            'entity_type': entity_type,
            'entity_id': entity_id,
            'exclude_actor': exclude_actor,
        }
    )


def notify_roles(
    roles,
    *,
    event_type,
    title,
    message='',
    link='',
    entity_type='',
    entity_id=None,
    actor=None,
    exclude_actor=True,
):
    """Wrapper calling generic notify_event engine for backward compatibility."""
    return notify_event(
        event_code=event_type,
        instance=None,
        actor=actor,
        extra_context={
            'explicit_roles': roles,
            'title': title,
            'message': message,
            'link': link,
            'entity_type': entity_type,
            'entity_id': entity_id,
            'exclude_actor': exclude_actor,
        }
    )


def notify_plate_replacement(plate_request, actor=None):
    return notify_event('plate.replacement_requested', instance=plate_request, actor=actor)


def notify_plate_cancelled(plate_request, actor=None):
    return notify_event('plate.cancelled', instance=plate_request, actor=actor)


def notify_sku_pending_review(recipe, actor=None):
    return notify_event('sku.pending_review', instance=recipe, actor=actor)


def notify_sku_sent_back(recipe, actor=None):
    return notify_event('sku.sent_back', instance=recipe, actor=actor)


def serialize_notification(item):
    return {
        'id': item.pk,
        'event_type': item.event_type,
        'title': item.title,
        'message': item.message,
        'link': item.link,
        'is_read': item.is_read,
        'created_at': item.created_at.isoformat() if item.created_at else '',
        'created_at_display': timezone.localtime(item.created_at).strftime('%Y-%m-%d %H:%M') if item.created_at else '',
    }


def log_rule_change(user, rule, action, old_db_obj=None):
    try:
        from core.models import NotificationRuleAuditLog
    except ImportError:
        return
    old_values = {}
    new_values = {}
    if action == 'update' and old_db_obj:
        for field in rule._meta.fields:
            field_name = field.name
            old_val = getattr(old_db_obj, field_name)
            new_val = getattr(rule, field_name)
            if old_val != new_val:
                old_values[field_name] = str(old_val.pk) if hasattr(old_val, 'pk') and old_val else str(old_val)
                new_values[field_name] = str(new_val.pk) if hasattr(new_val, 'pk') and new_val else str(new_val)
    elif action == 'create':
        for field in rule._meta.fields:
            field_name = field.name
            val = getattr(rule, field_name)
            new_values[field_name] = str(val.pk) if hasattr(val, 'pk') and val else str(val)
    elif action == 'delete':
        for field in rule._meta.fields:
            field_name = field.name
            val = getattr(rule, field_name)
            old_values[field_name] = str(val.pk) if hasattr(val, 'pk') and val else str(val)

    try:
        NotificationRuleAuditLog.objects.create(
            rule=rule if action != 'delete' else None,
            changed_by=user if getattr(user, 'pk', None) else None,
            action=action,
            old_values=old_values,
            new_values=new_values
        )
    except Exception as e:
        logger.warning("Error saving NotificationRuleAuditLog: %s", e)
