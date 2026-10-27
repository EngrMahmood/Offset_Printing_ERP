"""Flexo print-plan maths: order in PCS -> print in METERS (mirrors the
'Merge Sticker/Satin Planning' sheet).

Sheet formula:  Roll Qty Meter = (OrderQty - Stock) * MeterPerRoll / Ups
where OrderQty is in finished ROLLS and MeterPerRoll is the length of one finished roll.
The order unit is per SKU: label POs arrive in ROLLS (order_rolls = net qty), ribbon WOs in PIECES:

    order_rolls = net_qty                      (unit = roll)
    order_rolls = net_qty / pcs_per_roll       (unit = piece)
    base_m      = order_rolls * meter_per_roll / ups
    wastage_m   = base_m * wastage_pct / 100        (per-job %, 4.9 / 5 / 1.3 ... in the sheet)
    print_m     = base_m + wastage_m                (the sheet's "Actual Roll Qty Meter")

pcs_per_roll is entered per SKU (e.g. 5,000 labels per roll). When blank (e.g. ribbon with
no gap) it is derived from the piece length: pcs_per_roll = meter_per_roll * 1000 / size_h_mm.

Checks against the JC samples: 18,360 pcs, 5,000 pcs/roll, 142 m/roll, 11 ups ->
47.402 m + 5% = 49.772 m; 50,000 pcs x 304.8 mm, 200 m/roll, 1 up -> 15,240 m + 5% = 16,002 m.
"""
from dataclasses import dataclass
from decimal import Decimal

DEFAULT_WASTAGE_PCT = Decimal('5')
THOUSAND = Decimal('1000')


def _dec(value, default=None):
    if value is None or value == '':
        return default
    return value if isinstance(value, Decimal) else Decimal(str(value))


def effective_pcs_per_roll(*, pcs_per_roll=None, meter_per_roll=None, size_h_mm=None):
    """Finished pcs on one roll: entered (labels) or derived from piece length (ribbon)."""
    entered = _dec(pcs_per_roll, Decimal('0'))
    if entered > 0:
        return entered
    per_roll = _dec(meter_per_roll, Decimal('0'))
    length = _dec(size_h_mm, Decimal('0'))
    if per_roll > 0 and length > 0:
        return per_roll * THOUSAND / length
    return Decimal('0')


@dataclass(frozen=True)
class PrintPlan:
    net_qty: int
    order_rolls: Decimal
    base_meters: Decimal
    wastage_meters: Decimal
    print_meters: Decimal


def compute_print_plan(*, order_qty, stock_qty=0, ups, meter_per_roll, pcs_per_roll=None,
                       size_h_mm=None, wastage_pct=None, order_unit='piece'):
    order_qty = int(order_qty or 0)
    stock_qty = int(_dec(stock_qty, Decimal('0')))
    net_qty = max(order_qty - stock_qty, 0)

    ups = _dec(ups, Decimal('0'))
    per_roll_m = _dec(meter_per_roll, Decimal('0'))
    zero = Decimal('0')
    if order_unit == 'roll':
        per_roll_pcs = Decimal('1')
        order_rolls = Decimal(net_qty)
    else:
        per_roll_pcs = effective_pcs_per_roll(
            pcs_per_roll=pcs_per_roll, meter_per_roll=meter_per_roll, size_h_mm=size_h_mm)
        order_rolls = Decimal(net_qty) / per_roll_pcs if per_roll_pcs > 0 else zero
    if net_qty <= 0 or ups <= 0 or per_roll_m <= 0 or per_roll_pcs <= 0:
        return PrintPlan(net_qty, zero, zero, zero, zero)

    base = order_rolls * per_roll_m / ups
    wastage = base * _dec(wastage_pct, DEFAULT_WASTAGE_PCT) / Decimal('100')
    return PrintPlan(net_qty, order_rolls, base, wastage, base + wastage)


def pcs_per_print_meter(*, ups, meter_per_roll, pcs_per_roll=None, size_h_mm=None):
    """Meters -> pcs bridge for later production/WIP/dispatch phases (before wastage)."""
    per_roll_pcs = effective_pcs_per_roll(
        pcs_per_roll=pcs_per_roll, meter_per_roll=meter_per_roll, size_h_mm=size_h_mm)
    per_roll_m = _dec(meter_per_roll, Decimal('0'))
    ups = _dec(ups, Decimal('0'))
    if per_roll_m <= 0 or ups <= 0:
        return Decimal('0')
    return per_roll_pcs * ups / per_roll_m
