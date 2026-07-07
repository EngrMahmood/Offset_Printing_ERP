from __future__ import annotations

from reports.report_registry.metadata import ReportDefinition


class ReportRegistry:
    def __init__(self) -> None:
        self._reports: dict[str, ReportDefinition] = {}

    def register(self, definition: ReportDefinition) -> None:
        slug = (definition.slug or '').strip().lower()
        if not slug:
            raise ValueError('Report slug is required')
        if definition.executor is None:
            raise ValueError(f'Report {slug} requires an executor')
        self._reports[slug] = definition

    def get(self, slug: str) -> ReportDefinition | None:
        return self._reports.get((slug or '').strip().lower())

    def all(self) -> list[ReportDefinition]:
        return sorted(self._reports.values(), key=lambda item: (item.navigation_group, item.title))


registry = ReportRegistry()
