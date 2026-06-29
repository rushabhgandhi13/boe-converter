"""Property-based test for Totals aggregation (Value_Calculator.compute_totals).

Property 8: Column totals equal the sums of their per-line values, zero when empty.

**Validates: Requirements 7.1, 7.2, 7.3**

Strategy: generate lists of ``LineItem`` whose numeric inputs are sometimes
present (numeric) and sometimes missing/unparseable, run them through
``ValueCalculator.compute_line`` to obtain ``ComputedLine`` records, then assert
that every figure on the resulting ``Totals`` equals the order-preserving sum of
the corresponding per-line values (``None`` values skipped per Req 6.13). An
empty list must yield ``0`` for every total (Req 7.3).
"""

from __future__ import annotations

from hypothesis import given
from hypothesis import strategies as st

from boe_converter.calculator import ValueCalculator
from boe_converter.models import LineItem, RawValue

# Bounded, finite floats keep accumulated sums well clear of overflow/NaN so the
# aggregation can be compared for exact floating-point equality.
_FINITE = st.floats(
    min_value=-1e9,
    max_value=1e9,
    allow_nan=False,
    allow_infinity=False,
)


def _numeric_raw(value: float) -> RawValue:
    """A located, parseable numeric field as it would arrive from the parser."""

    return RawValue(raw_text=str(value), parsed=value)


# A field that is numeric, missing, or unparseable (exercise None-skipping).
_MAYBE_NUMERIC = st.one_of(
    _FINITE.map(_numeric_raw),
    st.just(RawValue.missing()),
    st.just(RawValue.unparseable("n/a")),
)

_UNIT = st.sampled_from(("DOZ", "PCS", "KGS", "SET")).map(
    lambda u: RawValue(raw_text=u, parsed=u)
)

_TEXT = RawValue(raw_text="item", parsed="item")

# A LineItem with independently varying (possibly absent) numeric inputs.
_LINE_ITEMS = st.builds(
    LineItem,
    item_serial=st.integers(min_value=1, max_value=10_000),
    cth_hsn=_FINITE.map(_numeric_raw),
    description=st.just(_TEXT),
    unit_price_usd=_MAYBE_NUMERIC,
    quantity=_MAYBE_NUMERIC,
    unit=_UNIT,
    assessable_value=_MAYBE_NUMERIC,
    bcd_rate=_FINITE.map(_numeric_raw),
    bcd_amount=_MAYBE_NUMERIC,
    igst_rate=_MAYBE_NUMERIC,
    total_duty=_FINITE.map(_numeric_raw),
)


def _sum_in_order(values) -> float:
    """Sum float|None values, skipping None, in iteration order (mirrors impl)."""

    total = 0.0
    for value in values:
        if value is not None:
            total += value
    return total


def _assessable_number(rv: RawValue) -> float | None:
    """The numeric assessable value, or None when missing/unparseable."""

    if rv.is_missing or rv.is_unparseable:
        return None
    parsed = rv.parsed
    if isinstance(parsed, bool) or not isinstance(parsed, (int, float)):
        return None
    return float(parsed)


@given(
    items=st.lists(_LINE_ITEMS, max_size=20),
    usd_rate=_FINITE,
    pkg_count=st.one_of(st.none(), st.integers(min_value=0, max_value=100_000)),
)
def test_totals_equal_sum_of_per_line_values(items, usd_rate, pkg_count):
    """Each total equals the order-preserving sum of its per-line values."""

    calc = ValueCalculator()
    lines = [calc.compute_line(item, usd_rate) for item in items]

    totals = calc.compute_totals(lines, pkg_count)

    assert totals.total_amount_usd == _sum_in_order(l.amount_usd for l in lines)
    assert totals.total_assessable_value == _sum_in_order(
        _assessable_number(l.source.assessable_value) for l in lines
    )
    assert totals.total_customs_duty == _sum_in_order(
        l.total_customs_duty for l in lines
    )
    assert totals.total_igst == _sum_in_order(l.igst_amount for l in lines)
    assert totals.total_land_cost_excl_gst == _sum_in_order(
        l.land_cost_excl_gst for l in lines
    )
    assert totals.total_land_cost_incl_gst == _sum_in_order(
        l.land_cost_incl_gst for l in lines
    )


@given(
    pkg_count=st.one_of(st.none(), st.integers(min_value=0, max_value=100_000)),
)
def test_totals_are_zero_when_no_lines(pkg_count):
    """An empty line list yields 0 for every column total (Req 7.3)."""

    totals = ValueCalculator().compute_totals([], pkg_count)

    assert totals.total_amount_usd == 0
    assert totals.total_assessable_value == 0
    assert totals.total_customs_duty == 0
    assert totals.total_igst == 0
    assert totals.total_land_cost_excl_gst == 0
    assert totals.total_land_cost_incl_gst == 0
