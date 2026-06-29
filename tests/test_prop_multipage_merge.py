"""Property 4: Multi-page item merge is equivalent to the unsplit item.

*For any* line item whose source block is split across a page boundary at an
arbitrary point, the merged record equals the record produced from the unsplit
block, with each field value appearing exactly once (no duplication, no omission).

This is modeled at the parser merge API level. The parser's
``_extract_invoice_items`` / ``_extract_duty_items`` collapse each item's
(possibly multi-page) block into a single ``InvoiceItemRow`` / ``DutyItemRow``
keyed by ``item_serial`` *before* the join in ``_merge_items``. This test
simulates that collapse:

1. A canonical item is generated, including a wrapped, multi-line description.
2. The item's source block is split across an arbitrary page boundary -- the
   description's line fragments are distributed between two pages, and the block
   is then stitched back into one row per serial exactly as the extract stage's
   collapse does (each fragment in order, each field once).
3. The ``LineItem`` produced by ``_merge_items`` from the stitched (multi-page)
   rows is asserted identical to the ``LineItem`` produced from the unsplit rows,
   with each field value appearing exactly once.

If the collapse/stitch ever duplicated a wrapped fragment or dropped one, the
exactly-once and equivalence assertions below would fail.

**Validates: Requirements 3.13**
"""

from __future__ import annotations

import hypothesis.strategies as st
from hypothesis import given

from boe_converter.models import RawValue
from boe_converter.parser import DutyItemRow, InvoiceItemRow, PdfParser

# Distinct, whitespace-free tokens for the wrapped description's line fragments,
# so each fragment can be counted unambiguously after stitching.
_TOKEN = st.text(
    alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", min_size=2, max_size=6
)


@st.composite
def _raw_scalar(draw):
    """Draw a scalar field captured verbatim as a ``RawValue``.

    Covers both a numeric value (the common Part II/III case) and short verbatim
    text, mirroring how ``_capture`` stores a field.
    """
    if draw(st.booleans()):
        value = draw(
            st.floats(
                min_value=0.0,
                max_value=1_000_000.0,
                allow_nan=False,
                allow_infinity=False,
            )
        )
        return RawValue(raw_text=repr(value), parsed=value)
    text = draw(st.text(min_size=1, max_size=8).filter(lambda s: s.strip()))
    return RawValue(raw_text=text, parsed=text)


@st.composite
def _multipage_item(draw):
    """Build a canonical item plus an arbitrary page split of its block.

    Returns ``(serial, fragments, split, inv_scalars, duty_scalars)`` where
    ``fragments`` are the (unique) line fragments of a wrapped description and
    ``split`` is the page-boundary index (``fragments[:split]`` land on the first
    page, ``fragments[split:]`` on the next).
    """
    serial = draw(st.integers(min_value=1, max_value=300))

    fragments = draw(st.lists(_TOKEN, min_size=1, max_size=6, unique=True))
    split = draw(st.integers(min_value=0, max_value=len(fragments)))

    inv_scalars = {
        name: draw(_raw_scalar())
        for name in ("cth_hsn", "unit_price_usd", "quantity", "unit", "amount")
    }
    duty_scalars = {
        name: draw(_raw_scalar())
        for name in (
            "assessable_value",
            "bcd_rate",
            "bcd_amount",
            "sws_amount",
            "igst_rate",
            "total_duty",
        )
    }
    return serial, fragments, split, inv_scalars, duty_scalars


def _build_rows(serial, description_text, inv_scalars, duty_scalars):
    """Build the ``(InvoiceItemRow, DutyItemRow)`` collapsed row pair for a serial."""
    description = RawValue(raw_text=description_text, parsed=description_text)
    inv = InvoiceItemRow(
        item_serial=serial,
        cth_hsn=inv_scalars["cth_hsn"],
        description=description,
        unit_price_usd=inv_scalars["unit_price_usd"],
        quantity=inv_scalars["quantity"],
        unit=inv_scalars["unit"],
        amount=inv_scalars["amount"],
    )
    duty = DutyItemRow(
        item_serial=serial,
        assessable_value=duty_scalars["assessable_value"],
        bcd_rate=duty_scalars["bcd_rate"],
        bcd_amount=duty_scalars["bcd_amount"],
        sws_amount=duty_scalars["sws_amount"],
        igst_rate=duty_scalars["igst_rate"],
        total_duty=duty_scalars["total_duty"],
    )
    return inv, duty


@given(item=_multipage_item())
def test_multipage_merge_equivalent_to_unsplit(item):
    serial, fragments, split, inv_scalars, duty_scalars = item
    parser = PdfParser()

    # Unsplit: the whole item sits on a single page -> description is the full,
    # in-order join of every wrapped fragment.
    unsplit_desc = " ".join(fragments)
    inv_u, duty_u = _build_rows(serial, unsplit_desc, inv_scalars, duty_scalars)

    # Multi-page: the same block is split across a page boundary at ``split`` and
    # then collapsed (stitched) back into one row per serial as the extract stage
    # does -> the wrapped description is the in-order concatenation of both page
    # fragments and every scalar field appears exactly once.
    page1_desc = fragments[:split]
    page2_desc = fragments[split:]
    stitched_desc = " ".join([*page1_desc, *page2_desc])
    inv_m, duty_m = _build_rows(serial, stitched_desc, inv_scalars, duty_scalars)

    items_u, flags_u = parser._merge_items(
        {serial: inv_u}, {serial: duty_u}, declared_count=None
    )
    items_m, flags_m = parser._merge_items(
        {serial: inv_m}, {serial: duty_m}, declared_count=None
    )

    # Exactly one record is produced from either form of the same item.
    assert len(items_u) == 1, "the unsplit item must yield exactly one record"
    assert len(items_m) == 1, "the multi-page item must yield exactly one record"
    line_u, line_m = items_u[0], items_m[0]

    # (a) Equivalence: the merged multi-page record equals the unsplit record,
    #     and neither produces spurious review flags.
    assert line_m == line_u, "multi-page merge must equal the unsplit merge"
    assert flags_m == flags_u == [], "a complete item must raise no review flags"

    # (b) Each wrapped description fragment appears exactly once -- no duplication
    #     across the page boundary and no omission.
    assert line_m.description.raw_text == unsplit_desc
    merged_tokens = line_m.description.raw_text.split()
    for fragment in fragments:
        assert merged_tokens.count(fragment) == 1, (
            f"fragment {fragment!r} must appear exactly once after stitching"
        )
    assert len(merged_tokens) == len(fragments), "no extra or missing fragments"

    # Every scalar field passes through unchanged, a single occurrence each.
    assert line_m.item_serial == serial
    assert line_m.cth_hsn == inv_scalars["cth_hsn"]
    assert line_m.unit_price_usd == inv_scalars["unit_price_usd"]
    assert line_m.quantity == inv_scalars["quantity"]
    assert line_m.unit == inv_scalars["unit"]
    assert line_m.assessable_value == duty_scalars["assessable_value"]
    assert line_m.bcd_rate == duty_scalars["bcd_rate"]
    assert line_m.bcd_amount == duty_scalars["bcd_amount"]
    assert line_m.igst_rate == duty_scalars["igst_rate"]
    assert line_m.total_duty == duty_scalars["total_duty"]
