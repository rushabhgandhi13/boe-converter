"""Property 7: missing/non-numeric computation inputs leave dependent values
blank and flagged.

*For any* line item in which a required computation input is missing or
non-numeric, every value that depends on that input is left blank (``None``,
no default substituted) and a ``ReviewFlag`` is emitted for the affected input.

**Validates: Requirements 6.13**
"""

from __future__ import annotations

import hypothesis.strategies as st
from hypothesis import HealthCheck, given, settings

from boe_converter.calculator import ValueCalculator
from boe_converter.models import LineItem, RawValue

# The required per-line numeric inputs to computation (Req 6.13). ``usd_rate`` is
# supplied as an argument and is always valid in this test, so the only missing
# inputs come from these line fields.
_REQUIRED_FIELDS = (
    "unit_price_usd",
    "quantity",
    "assessable_value",
    "bcd_amount",
    "igst_rate",
)

# For each computed ``ComputedLine`` field, the set of required inputs it depends
# on (derived from the normative formulas in design.md -> Computation model).
# ``purchase_rate_per_unit`` depends on quantity and land_cost_excl_gst; with a
# strictly-positive (non-zero) quantity its dependencies reduce to those below.
_DEPENDENCIES = {
    "amount_usd": {"unit_price_usd", "quantity"},
    "purchase_inr": {"unit_price_usd", "quantity"},
    "sws_amount": {"bcd_amount"},
    "total_customs_duty": {"bcd_amount"},
    "igst_amount": {"igst_rate", "assessable_value", "bcd_amount"},
    "combined_duty": {"bcd_amount", "igst_rate", "assessable_value"},
    "land_cost_excl_gst": {"unit_price_usd", "quantity", "bcd_amount"},
    "land_cost_incl_gst": {
        "unit_price_usd",
        "quantity",
        "bcd_amount",
        "igst_rate",
        "assessable_value",
    },
    "purchase_rate_per_unit": {"unit_price_usd", "quantity", "bcd_amount"},
}


def _valid_number():
    """Strictly-positive finite numbers, so no value is incidentally 0/NaN."""

    return st.floats(min_value=0.01, max_value=1_000_000.0, allow_nan=False, allow_infinity=False)


@st.composite
def _invalid_raw_value(draw):
    """A RawValue that is missing or non-numeric (the three Req 6.13 cases)."""

    kind = draw(st.sampled_from(("missing", "unparseable", "non_numeric_text")))
    if kind == "missing":
        return RawValue.missing()
    if kind == "unparseable":
        return RawValue.unparseable(draw(st.text(min_size=1, max_size=6)))
    # Located, not flagged unparseable, but parsed content is non-numeric text.
    text = draw(st.sampled_from(("N/A", "abc", "-", "  ", "USD")))
    return RawValue(raw_text=text, parsed=text)


@st.composite
def _line_with_missing(draw):
    """Build a LineItem with a non-empty random subset of required inputs invalid.

    Returns ``(item, invalid_fields)`` where ``invalid_fields`` is the set of
    required field names that were made missing/non-numeric.
    """

    # Choose which required fields are invalid; ensure at least one.
    invalid = draw(
        st.lists(st.sampled_from(_REQUIRED_FIELDS), min_size=1, unique=True).map(set)
    )

    fields = {}
    for name in _REQUIRED_FIELDS:
        if name in invalid:
            fields[name] = draw(_invalid_raw_value())
        else:
            value = draw(_valid_number())
            fields[name] = RawValue(raw_text=str(value), parsed=value)

    item = LineItem(
        item_serial=draw(st.integers(min_value=1, max_value=9999)),
        cth_hsn=RawValue(raw_text="1234", parsed="1234"),
        description=RawValue(raw_text="WIDGET", parsed="WIDGET"),
        unit_price_usd=fields["unit_price_usd"],
        quantity=fields["quantity"],
        unit=RawValue(raw_text="PCS", parsed="PCS"),  # non-DOZ: pcs stays blank
        assessable_value=fields["assessable_value"],
        bcd_rate=RawValue(raw_text="10", parsed=10.0),
        bcd_amount=fields["bcd_amount"],
        igst_rate=fields["igst_rate"],
        total_duty=RawValue(raw_text="0", parsed=0.0),
    )
    return item, invalid


@given(data=_line_with_missing(), usd_rate=_valid_number())
@settings(suppress_health_check=[HealthCheck.too_slow])
def test_missing_inputs_blank_dependent_values_and_flagged(data, usd_rate):
    item, invalid_fields = data
    calc = ValueCalculator()

    line = calc.compute_line(item, usd_rate)
    flags = calc.line_review_flags(item, usd_rate)

    # 1. A ReviewFlag is emitted for every affected (missing/non-numeric) input.
    flagged_fields = {f.field_name for f in flags if f.scope == "line_item"}
    for name in invalid_fields:
        assert name in flagged_fields, (
            f"expected a ReviewFlag for missing input {name!r}; got {flagged_fields}"
        )

    # 2. Every computed value depending on a missing input is left blank (None),
    #    with no default substituted.
    for computed_field, deps in _DEPENDENCIES.items():
        if deps & invalid_fields:
            assert getattr(line, computed_field) is None, (
                f"{computed_field} depends on missing input(s) "
                f"{deps & invalid_fields} but was not blank"
            )


@given(data=_line_with_missing(), usd_rate=_valid_number())
@settings(suppress_health_check=[HealthCheck.too_slow])
def test_no_spurious_flags_for_valid_inputs(data, usd_rate):
    """Inputs that are present and numeric are not flagged (no over-reporting)."""

    item, invalid_fields = data
    calc = ValueCalculator()

    flags = calc.line_review_flags(item, usd_rate)
    flagged_required = {
        f.field_name for f in flags if f.field_name in _REQUIRED_FIELDS
    }
    valid_fields = set(_REQUIRED_FIELDS) - invalid_fields
    assert flagged_required.isdisjoint(valid_fields), (
        f"valid inputs {valid_fields & flagged_required} were incorrectly flagged"
    )
