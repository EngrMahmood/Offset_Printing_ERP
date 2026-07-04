"""Shared plate-ink options for designer chips and production remake."""

PLATE_INK_OPTIONS = (
    'Cyan',
    'Magenta',
    'Yellow',
    'Black',
    'Special 1',
    'Special 2',
)

# Legacy chip labels still treated as inks (not production print color).
_PLATE_INK_ALIASES = {
    'special',
    'special1',
    'special2',
    'special 1',
    'special 2',
}


def is_plate_ink_spec(value):
    """True when value looks like plate ink chips (must never be stored as print color)."""
    parts = [part.strip().lower() for part in str(value or '').split(',') if part.strip()]
    if not parts:
        return False
    allowed = {name.lower() for name in PLATE_INK_OPTIONS} | _PLATE_INK_ALIASES
    return all(part in allowed for part in parts)
