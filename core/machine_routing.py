"""Colour/size based machine pool routing for the Machine Planning report.

Routing rules (see MACHINE_PLANNING_UPGRADE_PLAN.md, Part B2/B3):
1. Size gate first — a job whose sheet size exceeds the GTO groups' max
   print size is routed to the SM74 (largest sheet, 5-colour) machine
   regardless of colour count.
2. Otherwise colour decides the pool: "X+X" front/back is a single-colour
   job run in two passes; N (>=2) colours route to the double-colour pool
   expressed as passes = ceil(N / effective_colors).
3. A double-colour machine currently running on 1 operational colour drops
   into the single-colour pool for that period. A machine with
   operational_colors == 0 is excluded from all pools (under maintenance).
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field

MM_PER_INCH = 25.4

_SIZE_RE = re.compile(r'([0-9]+(?:\.[0-9]+)?)\s*[x*×]\s*([0-9]+(?:\.[0-9]+)?)', re.IGNORECASE)
_COLOR_PLUS_RE = re.compile(r'^\s*(\d+)\s*\+\s*(\d+)\s*$')
_COLOR_SINGLE_RE = re.compile(r'^\s*(\d+)\s*(?:colou?rs?)?\s*$', re.IGNORECASE)


def parse_sheet_size_mm(display):
    """Parse a free-text sheet size ('18*25', '520x740mm', ...) into (length_mm, width_mm).

    Values without a decimal point and both dimensions under 60 are assumed
    to be inches (this dataset stores press sheet sizes like '18*25' in
    inches); anything else is assumed to already be millimetres.
    """
    if not display:
        return None
    match = _SIZE_RE.search(str(display))
    if not match:
        return None
    a, b = float(match.group(1)), float(match.group(2))
    if a <= 60 and b <= 60:
        a, b = a * MM_PER_INCH, b * MM_PER_INCH
    return (a, b)


def color_class(color_spec_display):
    """Colour class used for machine routing.

    'X+X' (front/back of the same colour count) is a single-colour-class
    job run in two passes, so its class is X, not 2X. Asymmetric 'X+Y' and
    plain numeric specs fall back to their natural colour count.
    """
    raw = str(color_spec_display or '').strip()
    if not raw:
        return 0
    plus = _COLOR_PLUS_RE.match(raw)
    if plus:
        x, y = int(plus.group(1)), int(plus.group(2))
        return x if x == y else max(x, y)
    single = _COLOR_SINGLE_RE.match(raw)
    if single:
        return int(single.group(1))
    from core.print_colors import print_color_total_units
    return print_color_total_units(raw)


def _fits(size_mm, max_l, max_w):
    if size_mm is None or max_l is None or max_w is None:
        return True
    a, b = size_mm
    max_l, max_w = float(max_l), float(max_w)
    return (a <= max_l and b <= max_w) or (a <= max_w and b <= max_l)


@dataclass
class MachinePool:
    group_code: str
    members: list = field(default_factory=list)
    maintenance_members: list = field(default_factory=list)

    @property
    def label(self):
        names = [m.name for m in self.members]
        return ", ".join(names) if names else self.group_code

    @property
    def effective_colors(self):
        if not self.members:
            return 1
        return max(m.effective_colors for m in self.members)

    @property
    def max_size_mm(self):
        """(max_length, max_width) across pool members, or (None, None) if unset."""
        lengths = [m.max_print_length_mm for m in self.members if m.max_print_length_mm]
        widths = [m.max_print_width_mm for m in self.members if m.max_print_width_mm]
        return (max(lengths) if lengths else None, max(widths) if widths else None)


def pool_fits(pool, size_mm):
    """Whether a parsed (length_mm, width_mm) sheet size fits within a pool's max print size."""
    max_l, max_w = pool.max_size_mm
    return _fits(size_mm, max_l, max_w)


def build_pools(machines):
    """Group active machines by machine_group_code into GTO1/GTO2/SM74-style pools.

    A degraded double-colour machine (operational_colors == 1) is folded
    into the single-colour pool instead of its nominal group. A machine
    with operational_colors == 0 is excluded and tracked as maintenance.
    """
    pools = {}
    for m in machines:
        if getattr(m, 'machine_type', 'offset_printing') != 'offset_printing':
            continue
        code = (m.machine_group_code or '').strip()
        if not code:
            continue
        if m.is_under_maintenance:
            pools.setdefault(code, MachinePool(group_code=code)).maintenance_members.append(m)
            continue
        target_code = code
        if m.effective_colors == 1 and m.default_colors and m.default_colors > 1:
            target_code = _single_color_group_code(machines, fallback=code)
        pools.setdefault(target_code, MachinePool(group_code=target_code)).members.append(m)
    return pools


def _single_color_group_code(machines, fallback):
    """Best-effort group code for the fleet's single-colour pool."""
    candidates = [m.machine_group_code for m in machines if m.default_colors == 1 and m.machine_group_code]
    return candidates[0] if candidates else fallback


def find_pool_for_machine(pools, machine):
    """Return the MachinePool a specific machine belongs to (member or under
    maintenance), or None if it isn't part of any colour-grouped pool."""
    if machine is None:
        return None
    for pool in pools.values():
        if any(m.pk == machine.pk for m in pool.members) or any(m.pk == machine.pk for m in pool.maintenance_members):
            return pool
    return None


def _normalize_code(value):
    return re.sub(r'[^A-Z0-9]', '', str(value or '').upper())


def find_pool_by_group_code_text(pools, text):
    """Resolve a free-text machine assignment (e.g. 'GTO 1', legacy import
    data with no per-unit suffix) to its pool by matching machine_group_code,
    so it collapses into the same combined tab as 'GTO 1A' / 'GTO 1B' instead
    of showing up as a separate stray tab.
    """
    norm = _normalize_code(text)
    if not norm:
        return None
    for pool in pools.values():
        if any(_normalize_code(m.machine_group_code) == norm for m in pool.members):
            return pool
    return None


def route_job(color_spec_display, print_sheet_size_display, pools, size_gate_machine_code=None):
    """Resolve a job to a machine pool given its colour spec and sheet size.

    Returns a dict: pool_key, pool_label, member_machines, passes, color_class.
    Falls back to pool_key=None when there isn't enough data or no pools
    are configured, so callers can keep prior behaviour (e.g. literal
    machine_name) for jobs that can't be auto-routed yet.
    """
    if not pools:
        return None

    non_gate_pools = {k: v for k, v in pools.items() if k != size_gate_machine_code}
    gate_pool = pools.get(size_gate_machine_code) if size_gate_machine_code else None

    size_mm = parse_sheet_size_mm(print_sheet_size_display)
    if non_gate_pools:
        max_l = max((m.max_print_length_mm for p in non_gate_pools.values() for m in p.members if m.max_print_length_mm), default=None)
        max_w = max((m.max_print_width_mm for p in non_gate_pools.values() for m in p.members if m.max_print_width_mm), default=None)
    else:
        max_l = max_w = None

    if gate_pool and not _fits(size_mm, max_l, max_w):
        return {
            'pool_key': gate_pool.group_code,
            'pool_label': gate_pool.label,
            'member_machines': [m.name for m in gate_pool.members],
            'passes': 1,
            'color_class': color_class(color_spec_display),
        }

    colors = color_class(color_spec_display)
    if colors <= 0:
        return None

    # Prefer the smallest-capacity pool that can service this colour class in
    # a single pass; otherwise use the largest available pool and compute passes.
    ordered = sorted(non_gate_pools.values(), key=lambda p: p.effective_colors)
    target = None
    for pool in ordered:
        if colors <= pool.effective_colors:
            target = pool
            break
    if target is None:
        target = ordered[-1] if ordered else gate_pool

    if target is None:
        return None

    passes = max(1, math.ceil(colors / max(target.effective_colors, 1)))
    return {
        'pool_key': target.group_code,
        'pool_label': target.label,
        'member_machines': [m.name for m in target.members],
        'passes': passes,
        'color_class': colors,
    }
