from django import template

from coating import services

register = template.Library()


@register.inclusion_tag('coating/_panel.html')
def coating_panel(job_card):
    """Coating progress for a job card; renders nothing when the job has no coating and none was recorded."""
    if not job_card:
        return {'show': False}
    kind = services.coating_kind(services.job_application(job_card))
    info = services.summary(job_card)
    if not kind and not info['runs']:
        return {'show': False}
    return {'show': True, 'kind': kind, 'job_card': job_card, **info}
