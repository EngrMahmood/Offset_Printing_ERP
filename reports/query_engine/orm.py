from django.db.models import QuerySet


def optimize_queryset(queryset: QuerySet, *, select=(), prefetch=()):
    if select:
        queryset = queryset.select_related(*select)
    if prefetch:
        queryset = queryset.prefetch_related(*prefetch)
    return queryset
