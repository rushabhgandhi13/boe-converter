"""Smoke tests for the shared data models (task 1.1 scaffolding).

These verify the frozen dataclasses construct correctly, are immutable, and that
the ReviewFlagSet helper aggregates and queries flags as expected. Component
behavior tests live alongside their components in later tasks.
"""

from __future__ import annotations

import dataclasses

import pytest

from boe_converter.models import (
    ComputedDocument,
    ComputedLine,
    ConversionSummary,
    Discrepancy,
    ExtractedDocument,
    HeaderBlock,
    LineItem,
    RawValue,
    ReviewFlag,
    ReviewFlagSet,
    Totals,
)


def _raw(text: str, parsed=None) -> RawValue:
    return RawValue(raw_text=text, parsed=parsed if parsed is not None else text)


def _header() -> HeaderBlock:
    return HeaderBlock(
        company_name="Gemini Unicom LLP",
        party_name=_raw("ACME EXPORTS"),
        usd_rate=95.3,
        details=_raw("CO-32 CTN-1357"),
        invoice_no=_raw("INV-001"),
        invoice_date=_raw("01/06/2026"),
        be_no=_raw("2305090"),
        be_date=_raw("22/06/2026"),
        bl_no=_raw("BL-77"),
        bl_date=_raw("20/06/2026"),
        invoice_amount=_raw("12345.67", 12345.67),
        invoice_currency=_raw("USD"),
        package_count=_raw("1357", 1357),
        container_details=_raw("CONT-1 x 1"),
    )


def _line(serial: int = 1) -> LineItem:
    return LineItem(
        item_serial=serial,
        cth_hsn=_raw("39264049"),
        description=_raw("DECORATIVE ITEMS"),
        unit_price_usd=_raw("2.5", 2.5),
        quantity=_raw("100", 100.0),
        unit=_raw("DOZ"),
        assessable_value=_raw("23825.0", 23825.0),
        bcd_rate=_raw("10", 10.0),
        bcd_amount=_raw("2382.5", 2382.5),
        igst_rate=_raw("0.18", 0.18),
        total_duty=_raw("5000.0", 5000.0),
    )


def test_raw_value_defaults_and_helpers():
    rv = RawValue()
    assert rv.raw_text is None and rv.parsed is None
    assert rv.is_missing is False and rv.is_unparseable is False

    missing = RawValue.missing()
    assert missing.is_missing is True
    assert missing.raw_text is None

    bad = RawValue.unparseable("12.x.3")
    assert bad.is_unparseable is True
    assert bad.raw_text == "12.x.3"


def test_models_are_frozen():
    rv = RawValue(raw_text="x")
    with pytest.raises(dataclasses.FrozenInstanceError):
        rv.raw_text = "y"  # type: ignore[misc]


def test_construct_extracted_document():
    doc = ExtractedDocument(
        header=_header(),
        line_items=[_line(1), _line(2)],
        declared_item_count=2,
    )
    assert len(doc.line_items) == 2
    assert doc.declared_item_count == 2
    assert doc.flags == []


def test_construct_computed_document():
    line = ComputedLine(source=_line(1), amount_usd=250.0)
    totals = Totals(total_amount_usd=250.0, package_count=_raw("1357", 1357))
    doc = ComputedDocument(header=_header(), lines=[line], totals=totals)
    assert doc.lines[0].amount_usd == 250.0
    assert doc.totals.total_amount_usd == 250.0


def test_totals_defaults_to_zero():
    t = Totals()
    assert t.total_amount_usd == 0.0
    assert t.total_igst == 0.0
    assert isinstance(t.package_count, RawValue)


def test_conversion_summary_construction():
    flag = ReviewFlag(scope="line_item", field_name="quantity", reason="MISSING", item_serial=3)
    disc = Discrepancy(kind="ITEM_COUNT", message="mismatch", expected=45, actual=44)
    summary = ConversionSummary(
        line_items_extracted=44,
        declared_item_count=45,
        total_invoice_amount_usd=1000.0,
        declared_invoice_amount_usd=1000.0,
        review_flag_count=1,
        review_flags=[flag],
        discrepancies=[disc],
    )
    assert summary.review_flag_count == 1
    assert summary.discrepancies[0].kind == "ITEM_COUNT"


def test_review_flag_set_aggregation_and_queries():
    fs = ReviewFlagSet()
    fs.add(ReviewFlag(scope="header", field_name="be_no", reason="MISSING"))
    fs.extend(
        [
            ReviewFlag(scope="line_item", field_name="quantity", reason="MISSING", item_serial=2),
            ReviewFlag(
                scope="line_item",
                field_name="unit_price_usd",
                reason="UNPARSEABLE",
                item_serial=2,
                raw_text="2.x5",
            ),
        ]
    )

    assert fs.count() == 3
    assert len(fs) == 3
    assert fs.is_header_flagged("be_no") is True
    assert fs.is_header_flagged("invoice_no") is False
    assert fs.is_line_flagged(2, "quantity") is True
    assert fs.is_line_flagged(2, "assessable_value") is False
    assert len(fs.for_line(2)) == 2
    # flags property returns a copy
    snapshot = fs.flags
    snapshot.clear()
    assert fs.count() == 3


def test_review_flag_set_iteration():
    flags = [
        ReviewFlag(scope="totals", field_name="package_count", reason="MISSING"),
    ]
    fs = ReviewFlagSet(flags)
    collected = list(fs)
    assert collected == flags
