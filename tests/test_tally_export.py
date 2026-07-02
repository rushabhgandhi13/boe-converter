"""Tests for the Tally master matching and Purchase-voucher JSON export.

These use a small **synthetic** in-memory master (no commercial data) so they
run in CI without the real ``Master.json``.
"""

from __future__ import annotations

import json

import pytest
from hypothesis import given, settings
from hypothesis import strategies as st

from boe_converter.excel_reader import ExcelReadError, read_workbook
from boe_converter.excel_writer import ExcelGenerator
from boe_converter.models import (
    ComputedDocument,
    ComputedLine,
    HeaderBlock,
    LineItem,
    RawValue,
    ReviewFlagSet,
    Totals,
)
from boe_converter.tally_exporter import CompanyProfile, TallyExporter
from boe_converter.tally_master import TallyMaster


# ---------------------------------------------------------------------------
# Synthetic master + document builders
# ---------------------------------------------------------------------------
def _ledger(name: str) -> dict:
    return {"metadata": {"type": "Ledger", "name": name}}


def make_master(extra: list[str] | None = None) -> TallyMaster:
    names = [
        "Factory Purchase (Import 5%)",
        "Factory Purchase (Import 18%)",
        "IGST Purchase @ 5.00 %",
        "IGST Purchase @ 18%",
        "IGST Payable @ 5%",
        "IGST Payable @ 18%",
        "Custom Duty Payable",
        "Tax Free (Purchases)",
        "Prayan Impex Company Limited",
    ]
    names += extra or []
    doc = {"tallymessage": [_ledger(n) for n in names]}
    return TallyMaster(doc)


def rv(value) -> RawValue:
    return RawValue(raw_text=str(value), parsed=value)


def make_line(serial: int, rate: float, qty: float, unit: str, hsn: str) -> ComputedLine:
    purchase_inr = 1000.0 * serial
    duty = 100.0 * serial
    land_excl = purchase_inr + duty
    igst = round(rate * (purchase_inr + duty), 2)
    src = LineItem(
        item_serial=serial,
        cth_hsn=rv(hsn),
        description=rv(f"ITEM {serial}"),
        unit_price_usd=rv(10.0),
        quantity=rv(qty),
        unit=rv(unit),
        assessable_value=rv(purchase_inr),
        bcd_rate=rv(0.1),
        bcd_amount=rv(duty),
        igst_rate=rv(rate),
        total_duty=rv(duty + igst),
    )
    return ComputedLine(
        source=src,
        amount_usd=10.0 * qty,
        purchase_inr=purchase_inr,
        total_customs_duty=duty,
        igst_amount=igst,
        land_cost_excl_gst=land_excl,
        land_cost_incl_gst=land_excl + igst,
        purchase_rate_per_unit=land_excl / qty if qty else 0.0,
    )


def make_doc(lines: list[ComputedLine]) -> ComputedDocument:
    header = HeaderBlock(
        company_name="Gemini Unicom LLP",
        party_name=rv("PRAYAN IMPEX COMPANY LIMITED"),
        usd_rate=95.3,
        details=rv("CO-26 CTN-908"),
        invoice_no=rv("ZJXY26050983"),
        invoice_date=rv("01/06/2026"),
        be_no=rv("2022823"),
        be_date=rv("27/06/2026"),
        bl_no=RawValue.missing(),
        bl_date=RawValue.missing(),
        invoice_amount=rv(sum(l.amount_usd or 0 for l in lines)),
        invoice_currency=rv("USD"),
        package_count=rv(908),
        container_details=RawValue.missing(),
    )
    totals = Totals(
        total_amount_usd=sum(l.amount_usd or 0 for l in lines),
        total_assessable_value=sum(l.purchase_inr or 0 for l in lines),
        total_customs_duty=sum(l.total_customs_duty or 0 for l in lines),
        total_igst=sum(l.igst_amount or 0 for l in lines),
        total_land_cost_excl_gst=sum(l.land_cost_excl_gst or 0 for l in lines),
        total_land_cost_incl_gst=sum(l.land_cost_incl_gst or 0 for l in lines),
    )
    return ComputedDocument(header=header, lines=lines, totals=totals, flags=[])


# ---------------------------------------------------------------------------
# Master matching
# ---------------------------------------------------------------------------
def test_match_uses_canonical_master_name():
    m = make_master()
    # Query in uppercase should resolve to the master's title-case spelling.
    assert m.resolve("Prayan Impex Company Limited") == "Prayan Impex Company Limited"
    res = m.match("PRAYAN IMPEX COMPANY LIMITED")
    assert res.matched and res.name == "Prayan Impex Company Limited"


def test_match_rate_ledger_picks_rate_specific_name():
    m = make_master()
    assert m.match_rate_ledger("IGST Purchase", 0.05).name == "IGST Purchase @ 5.00 %"
    assert m.match_rate_ledger("IGST Purchase", 0.18).name == "IGST Purchase @ 18%"
    assert m.match_rate_ledger("Factory Purchase Import", 0.05).name == "Factory Purchase (Import 5%)"


def test_missing_detects_unknown_supplier():
    m = make_master()
    missing = m.missing(["Prayan Impex Company Limited", "Totally New Supplier Ltd"])
    assert missing == ["Totally New Supplier Ltd"]


def test_add_ledger_then_matches_and_serialises_utf16():
    m = make_master()
    m.add_ledger("Totally New Supplier Ltd")
    assert m.has_ledger("Totally New Supplier Ltd")
    raw = m.to_bytes()
    # Round-trips as UTF-16 and the new ledger is present.
    reloaded = TallyMaster.load_bytes(raw)
    assert reloaded.has_ledger("Totally New Supplier Ltd")


# ---------------------------------------------------------------------------
# Voucher construction
# ---------------------------------------------------------------------------
def _signed_total(entries: list[dict]) -> float:
    """Sum ledger amounts with credit(+)/debit(-) sign for balance checking."""
    total = 0.0
    for e in entries:
        total += float(e["amount"])
    return total


def test_voucher_balances_to_zero():
    m = make_master()
    doc = make_doc([
        make_line(1, 0.05, 100, "KGS", "96071190"),
        make_line(2, 0.18, 50, "PCS", "39264049"),
    ])
    exporter = TallyExporter(m, CompanyProfile(name="Gemini Unicom LLP", gstin="27AAA0000A1Z0"))
    out = exporter.build(doc, usd_rate=95.3)
    entries = out["tallymessage"][0]["allledgerentries"]
    assert abs(_signed_total(entries)) < 0.01


def test_voucher_groups_ledgers_by_rate():
    m = make_master()
    doc = make_doc([
        make_line(1, 0.05, 100, "KGS", "96071190"),
        make_line(2, 0.18, 50, "PCS", "39264049"),
        make_line(3, 0.0, 10, "NOS", "10000000"),
    ])
    exporter = TallyExporter(m)
    entries = exporter.build(doc, 95.3)["tallymessage"][0]["allledgerentries"]
    names = [e["ledgername"] for e in entries]
    assert "Factory Purchase (Import 5%)" in names
    assert "Factory Purchase (Import 18%)" in names
    assert "Tax Free (Purchases)" in names
    assert "Custom Duty Payable" in names
    assert "IGST Purchase @ 5.00 %" in names
    assert "IGST Payable @ 18%" in names


def test_party_amount_equals_total_purchase_inr():
    m = make_master()
    lines = [make_line(1, 0.05, 100, "KGS", "96071190"), make_line(2, 0.18, 50, "PCS", "39264049")]
    doc = make_doc(lines)
    entries = TallyExporter(m).build(doc, 95.3)["tallymessage"][0]["allledgerentries"]
    party = next(e for e in entries if e.get("ispartyledger"))
    expected = sum(l.purchase_inr for l in lines)
    assert abs(float(party["amount"]) - expected) < 0.01


def test_inventory_allocations_present_and_named():
    m = make_master()
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "96071190")])
    entries = TallyExporter(m).build(doc, 95.3)["tallymessage"][0]["allledgerentries"]
    purchase = next(e for e in entries if e["ledgername"] == "Factory Purchase (Import 5%)")
    inv = purchase["inventoryallocations"]
    assert len(inv) == 1
    assert inv[0]["stockitemname"] == "ITEM 1"
    assert inv[0]["gsthsnname"] == "96071190"


def test_custom_duty_equals_total_duty():
    m = make_master()
    lines = [make_line(1, 0.05, 100, "KGS", "9"), make_line(2, 0.18, 50, "PCS", "3")]
    doc = make_doc(lines)
    entries = TallyExporter(m).build(doc, 95.3)["tallymessage"][0]["allledgerentries"]
    duty = next(e for e in entries if e["ledgername"] == "Custom Duty Payable")
    assert abs(float(duty["amount"]) - sum(l.total_customs_duty for l in lines)) < 0.01


def test_required_ledger_names_include_all_groups():
    m = make_master()
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "9"), make_line(2, 0.18, 50, "PCS", "3")])
    names = TallyExporter(m).required_ledger_names(doc)
    assert "Custom Duty Payable" in names
    assert "Prayan Impex Company Limited" in names


def test_output_is_json_serialisable():
    m = make_master()
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "9")])
    out = TallyExporter(m).build(doc, 95.3)
    # Must serialise cleanly for download.
    json.dumps(out)


# ---------------------------------------------------------------------------
# Excel round-trip (manual-upload path)
# ---------------------------------------------------------------------------
def test_excel_reader_reconstructs_lines_without_formulas():
    m = make_master()
    lines = [make_line(1, 0.05, 100, "KGS", "96071190"), make_line(2, 0.18, 50, "PCS", "39264049")]
    doc = make_doc(lines)
    # Generate with numeric values (not formulas) so cached results exist.
    gen = ExcelGenerator(use_formulas=False)
    xlsx = gen.generate(doc, ReviewFlagSet([]))
    rebuilt = read_workbook(xlsx)
    assert len(rebuilt.lines) == len(lines)
    # A voucher built from the re-read document should still balance.
    entries = TallyExporter(m).build(rebuilt, 95.3)["tallymessage"][0]["allledgerentries"]
    assert abs(_signed_total(entries)) < 0.5


def test_excel_reader_rejects_empty_workbook():
    from openpyxl import Workbook

    import io

    wb = Workbook()
    buf = io.BytesIO()
    wb.save(buf)
    with pytest.raises(ExcelReadError):
        read_workbook(buf.getvalue())


# ---------------------------------------------------------------------------
# Property: any mix of rates/lines produces a balanced voucher
# ---------------------------------------------------------------------------
@settings(max_examples=60, deadline=None)
@given(
    specs=st.lists(
        st.tuples(
            st.sampled_from([0.0, 0.05, 0.12, 0.18, 0.28]),
            st.integers(min_value=1, max_value=500),
        ),
        min_size=1,
        max_size=12,
    )
)
def test_property_voucher_always_balances(specs):
    extra = [
        "Factory Purchase (Import 12%)",
        "Factory Purchase (Import 28%)",
        "IGST Purchase @ 12%",
        "IGST Purchase @ 28%",
        "IGST Payable @ 12%",
        "IGST Payable @ 28%",
    ]
    m = make_master(extra)
    lines = [
        make_line(i + 1, rate, qty, "NOS", "1000")
        for i, (rate, qty) in enumerate(specs)
    ]
    doc = make_doc(lines)
    entries = TallyExporter(m).build(doc, 95.3)["tallymessage"][0]["allledgerentries"]
    assert abs(_signed_total(entries)) < 0.05
