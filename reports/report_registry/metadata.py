from dataclasses import dataclass, field
from typing import Callable


@dataclass(frozen=True)
class ReportDefinition:
    slug: str
    title: str
    description: str
    department: str
    permissions: tuple[str, ...] = field(default_factory=tuple)
    filters: tuple[str, ...] = field(default_factory=tuple)
    supported_exports: tuple[str, ...] = field(default_factory=tuple)
    supported_charts: tuple[str, ...] = field(default_factory=tuple)
    drilldown_support: bool = False
    cache_timeout: int = 300
    icon: str = ''
    category: str = ''
    navigation_group: str = ''
    executor: Callable | None = None
