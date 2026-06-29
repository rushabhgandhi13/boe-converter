"""Property 1: No line item is ever dropped, and missing/failed fields are flagged.

*For any* collection of line items where an arbitrary subset of required fields is
missing or fails to parse, every such line item still appears in the merged output
exactly once (in ascending serial order), and each missing or unparseable required
field produces a ``ReviewFlag`` identifying that line item and field, with no
inferred or default value substituted.

This exercises ``PdfParser._merge_items`` directly: it joins Part II invoice rows
(``InvoiceItemRow``) and Part III duty rows (``DutyItemRow``) keyed by
``item_serial``. The strategy generates dicts with overlapping *and* disjoint serial
sets, where some fields are missing/unparseable, so the no-drop and flagging
invariants are checked across the full join space.

**Validates: Requirements 3.1, 3.15, 5.11, 9.5, 2.10**
"""

from __future__ import annotations

import hypothesis.strategies as st
from hypothesis import given

from boe_converter.models import RawValue
from boe_converter.parser import DutyItemRow, InvoiceItemRow, PdfParser

# Required line-item fields sourced from Part II (invoice) and Part III (duty).
# These mirror the join in ``_merge_items`` (Req 3.4-3.12; item_serial is the
# always-present int join key and so is not a flaggable field).
_INVOICE_FIELDS = ("cth_hsn", "description", "unit_price_usd", "quantity", "unit")
_DUTY_FIELDS = (
    "assessable_value",
    "bcd_rate",
    "bcd_amount",
    "igst_rate",
    "total_duty",
)


@st.composite
def _field_value(draw):
    """Draw a ``RawValue`` and the review reason it should produce, if any.

    Covers the three states the merge distinguishes per field:
    - valid (numeric or plain text): no flag (``reason is None``)
    - missing: a ``MISSING`` flag
    - unparseable: an ``UNPARSEABLE`` flag (raw text retained)
    """
    kind = draw(st.sampled_from(("valid_num", "valid_text", "missing", "unparseable")))
    if kind == "missing":
        return RawValue.missing(), "MISSING"
    if kind == "unparseable":
        return RawValue.unparseable(draw(st.text(min_size=1, max_size=6))), "UNPARSEABLE"
    if kind == "valid_num":
        value = draw(
            st.floats(min_value=0.0, max_value=1_000_000.0, allow_nan=False, allow_infinity=False)
        )
        return RawValue(raw_text=str(value), parsed=value), None
    text = draw(st.text(min_size=1, max_size=6).filter(lambda s: s.strip()))
    return RawValue(raw_text=text, parsed=text), None


@st.composite
def _merge_scenario(draw):
    """Build (inv, duty, all_serials, expected_flags) for ``_merge_items``.

    Each serial is independently assigned to Part II only, Part III only, or both,
    so the generated dicts span overlapping and disjoint serial sets. A serial
    present in only one source has every field of the *absent* source treated as
    missing (and therefore flagged ``MISSING``), per Req 3.1/9.5.

    ``expected_flags`` is the exact set of ``(item_serial, field_name, reason)``
    tuples the merge must emit for ``line_item`` scope.
    """
    serials = draw(
        st.lists(st.integers(min_value=1, max_value=300), min_size=0, max_size=8, unique=True)
    )

    inv: dict[int, InvoiceItemRow] = {}
    duty: dict[int, DutyItemRow] = {}
    expected_flags: set[tuple[int, str, str]] = set()

    for serial in serials:
        category = draw(st.sampled_from(("inv", "duty", "both")))

        # --- Part II (invoice) ---
        if category in ("inv", "both"):
            inv_vals: dict[str, RawValue] = {}
            for name in _INVOICE_FIELDS:
                rv, reason = draw(_field_value())
                inv_vals[name] = rv
                if reason is not None:
                    expected_flags.add((serial, name, reason))
            inv[serial] = InvoiceItemRow(
                item_serial=serial,
                cth_hsn=inv_vals["cth_hsn"],
                description=inv_vals["description"],
                unit_price_usd=inv_vals["unit_price_usd"],
                quantity=inv_vals["quantity"],
                unit=inv_vals["unit"],
                # amount is not a merged required field; keep it valid.
                amount=RawValue(raw_text="1", parsed=1.0),
            )
        else:
            # Invoice source absent -> all invoice fields missing on the LineItem.
            for name in _INVOICE_FIELDS:
                expected_flags.add((serial, name, "MISSING"))

        # --- Part III (duty) ---
        if category in ("duty", "both"):
            duty_vals: dict[str, RawValue] = {}
            for name in _DUTY_FIELDS:
                rv, reason = draw(_field_value())
                duty_vals[name] = rv
                if reason is not None:
                    expected_flags.add((serial, name, reason))
            duty[serial] = DutyItemRow(
                item_serial=serial,
                assessable_value=duty_vals["assessable_value"],
                bcd_rate=duty_vals["bcd_rate"],
                bcd_amount=duty_vals["bcd_amount"],
                # sws_amount is not a merged required field; keep it valid.
                sws_amount=RawValue(raw_text="0", parsed=0.0),
                igst_rate=duty_vals["igst_rate"],
                total_duty=duty_vals["total_duty"],
            )
        else:
            # Duty source absent -> all duty fields missing on the LineItem.
            for name in _DUTY_FIELDS:
                expected_flags.add((serial, name, "MISSING"))

    return inv, duty, set(serials), expected_flags


@given(scenario=_merge_scenario())
def test_no_line_item_dropped_and_missing_fields_flagged(scenario):
    inv, duty, all_serials, expected_flags = scenario
    parser = PdfParser()

    line_items, flags = parser._merge_items(inv, duty, declared_count=None)

    # (a) Every serial present in either source appears exactly once, ascending.
    result_serials = [li.item_serial for li in line_items]
    assert result_serials == sorted(all_serials), (
        "merged serials must be every input serial, each once, in ascending order"
    )
    assert len(result_serials) == len(set(result_serials)), "no duplicate line items"

    # (b) Every missing/unparseable required field produces exactly the expected
    #     ReviewFlag for that serial -- no drops, no spurious flags, no defaults.
    line_flags = [f for f in flags if f.scope == "line_item"]
    actual_flags = {(f.item_serial, f.field_name, f.reason) for f in line_flags}
    assert actual_flags == expected_flags
    assert len(line_flags) == len(expected_flags), "each flag emitted exactly once"
