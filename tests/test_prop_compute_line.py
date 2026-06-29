"""Property-based test for per-line monetary computations (task 2.2).

Property 5: Per-line monetary computations satisfy their formulas and algebraic
relations.

**Validates: Requirements 6.1, 6.2, 6.3, 6.4, 6.5, 6.6, 6.7, 6.8, 6.9, 6.10**

For any line item with numeric inputs (unit price, quantity, USD rate,
assessable value, BCD amount, IGST rate), ``ValueCalculator.compute_line``
reproduces each normative formula from design.md -> Computation model and the
algebraic relations between the derived values:

    amount_usd             = unit_price * qty                       (6.1)
    purchase_inr           = amount_usd * usd_rate                  (6.2)
    sws_amount             = bcd_amount * 0.10                      (6.3)
    total_customs_duty     = bcd_amount + sws_amount               (6.4)
    igst_amount            = igst_rate * (assess + total_customs_duty)  (6.5)
    combined_duty          = total_customs_duty + igst_amount       (6.6)
    land_cost_excl_gst     = purchase_inr + total_customs_duty      (6.7)
    land_cost_incl_gst     = land_cost_excl_gst + igst_amount       (6.8)
    purchase_rate_per_unit = land_cost_excl_gst / qty               (6.9)
                           = 0 when qty == 0 (no division)          (6.10)

Computations retain full floating-point precision and the calculator performs
exactly these operations, so the reproduced formulas are asserted with exact
equality. The independent algebraic relations are likewise exact because they
reuse the same intermediate float values.
"""

from __future__ import annotations

from hypothesis import given
from hypothesis import strategies as st

from boe_converter.calculator import ValueCalculator
from boe_converter.models import LineItem, RawValue

SWS_RATE = 0.10


def _raw_num(value: float) -> RawValue:
    """A RawValue carrying a numeric ``parsed`` interpretation, as the parser emits."""

    return RawValue(raw_text=repr(value), parsed=value)


# Finite numeric inputs kept to moderate magnitudes so chained products stay
# well within float range (no inf/overflow), while still exercising negatives,
# zero, and fractional values across the input space.
finite_money = st.floats(
    min_value=-1e6,
    max_value=1e6,
    allow_nan=False,
    allow_infinity=False,
)
quantities = st.floats(
    min_value=-1e5,
    max_value=1e5,
    allow_nan=False,
    allow_infinity=False,
)
igst_rates = st.floats(
    min_value=0.0,
    max_value=1.0,
    allow_nan=False,
    allow_infinity=False,
)
usd_rates = st.floats(
    min_value=0.0,
    max_value=200.0,
    allow_nan=False,
    allow_infinity=False,
)


@st.composite
def numeric_line_items(draw):
    """Build a LineItem whose five required numeric inputs all parse cleanly."""

    item = LineItem(
        item_serial=draw(st.integers(min_value=1, max_value=10_000)),
        cth_hsn=RawValue(raw_text="39264049", parsed="39264049"),
        description=RawValue(raw_text="ITEM", parsed="ITEM"),
        unit_price_usd=_raw_num(draw(finite_money)),
        quantity=_raw_num(draw(quantities)),
        unit=RawValue(raw_text="DOZ", parsed="DOZ"),
        assessable_value=_raw_num(draw(finite_money)),
        bcd_rate=_raw_num(draw(finite_money)),
        bcd_amount=_raw_num(draw(finite_money)),
        igst_rate=_raw_num(draw(igst_rates)),
        total_duty=_raw_num(draw(finite_money)),
    )
    return item


@given(item=numeric_line_items(), usd_rate=usd_rates)
def test_per_line_formulas_and_relations(item: LineItem, usd_rate: float):
    """compute_line reproduces every normative formula and algebraic relation."""

    unit_price = float(item.unit_price_usd.parsed)
    qty = float(item.quantity.parsed)
    assessable = float(item.assessable_value.parsed)
    bcd_amount = float(item.bcd_amount.parsed)
    igst_rate = float(item.igst_rate.parsed)

    line = ValueCalculator().compute_line(item, usd_rate)

    # 6.1 amount_usd = unit_price * qty
    expected_amount_usd = unit_price * qty
    assert line.amount_usd == expected_amount_usd

    # 6.2 purchase_inr = amount_usd * usd_rate
    expected_purchase_inr = expected_amount_usd * usd_rate
    assert line.purchase_inr == expected_purchase_inr

    # 6.3 sws_amount = bcd_amount * 0.10
    expected_sws = bcd_amount * SWS_RATE
    assert line.sws_amount == expected_sws

    # 6.4 total_customs_duty = bcd_amount + sws_amount
    expected_total_customs_duty = bcd_amount + expected_sws
    assert line.total_customs_duty == expected_total_customs_duty

    # 6.5 igst_amount = igst_rate * (assessable + total_customs_duty)
    expected_igst = igst_rate * (assessable + expected_total_customs_duty)
    assert line.igst_amount == expected_igst

    # 6.6 combined_duty = total_customs_duty + igst_amount
    expected_combined = expected_total_customs_duty + expected_igst
    assert line.combined_duty == expected_combined

    # 6.7 land_cost_excl_gst = purchase_inr + total_customs_duty
    expected_land_excl = expected_purchase_inr + expected_total_customs_duty
    assert line.land_cost_excl_gst == expected_land_excl

    # 6.8 land_cost_incl_gst = land_cost_excl_gst + igst_amount
    expected_land_incl = expected_land_excl + expected_igst
    assert line.land_cost_incl_gst == expected_land_incl

    # 6.9 / 6.10 purchase_rate_per_unit = land_cost_excl_gst / qty (0 when qty==0)
    if qty == 0:
        assert line.purchase_rate_per_unit == 0
    else:
        assert line.purchase_rate_per_unit == expected_land_excl / qty

    # Algebraic relations stated directly against the computed fields.
    assert line.combined_duty == line.total_customs_duty + line.igst_amount
    assert line.land_cost_incl_gst == line.land_cost_excl_gst + line.igst_amount
    assert line.total_customs_duty == line.sws_amount + bcd_amount
    assert line.land_cost_excl_gst == line.purchase_inr + line.total_customs_duty
