"""Tests for the Tally Purchase-voucher JSON export (Master.json-free).

Ledger names follow Tally's deterministic conventions and buyer/seller identity
comes from the BOE ``HeaderBlock`` (with optional profile overrides). All data
here is synthetic so the suite runs in CI without any commercial files.
"""

from __future__ import annotations

import io
import json

import pytest
from hypothesis import given, settings
from hypothesis import strategies as st

from boe_converter.excel_reader import ExcelReadError, read_workbook, write_tally_names
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
from boe_converter.tally_exporter import (
    CompanyProfile,
    SellerProfile,
    TallyExporter,
    apply_stock_names,
)


# ---------------------------------------------------------------------------
# Synthetic document builders
# ---------------------------------------------------------------------------
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


def make_doc(lines: list[ComputedLine], **header_over) -> ComputedDocument:
    base = dict(
        company_name="M/S GEMINI UNICOM LLP",
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
        buyer_gstin=rv("27AAYFG7003K1ZW"),
        buyer_address=rv("A/43, DEVIDAYAL ROAD, MULUND"),
        buyer_pincode=rv("400080"),
        buyer_state=rv("Maharashtra"),
        seller_address=rv("ROOM 203, YIWU"),
        seller_country=rv("China"),
    )
    base.update(header_over)
    header = HeaderBlock(**base)
    totals = Totals(
        total_amount_usd=sum(l.amount_usd or 0 for l in lines),
        total_assessable_value=sum(l.purchase_inr or 0 for l in lines),
        total_customs_duty=sum(l.total_customs_duty or 0 for l in lines),
        total_igst=sum(l.igst_amount or 0 for l in lines),
        total_land_cost_excl_gst=sum(l.land_cost_excl_gst or 0 for l in lines),
        total_land_cost_incl_gst=sum(l.land_cost_incl_gst or 0 for l in lines),
    )
    return ComputedDocument(header=header, lines=lines, totals=totals, flags=[])


def _signed_total(entries: list[dict]) -> float:
    return sum(float(e["amount"]) for e in entries)


# ---------------------------------------------------------------------------
# Ledger naming + voucher construction
# ---------------------------------------------------------------------------
def test_voucher_balances_to_zero():
    doc = make_doc([
        make_line(1, 0.05, 100, "KGS", "96071190"),
        make_line(2, 0.18, 50, "PCS", "39264049"),
    ])
    out = TallyExporter().build(doc, usd_rate=95.3)
    entries = out["tallymessage"][0]["allledgerentries"]
    assert abs(_signed_total(entries)) < 0.01


def test_voucher_groups_ledgers_by_rate():
    doc = make_doc([
        make_line(1, 0.05, 100, "KGS", "96071190"),
        make_line(2, 0.18, 50, "PCS", "39264049"),
        make_line(3, 0.0, 10, "NOS", "10000000"),
    ])
    entries = TallyExporter().build(doc, 95.3)["tallymessage"][0]["allledgerentries"]
    names = [e["ledgername"] for e in entries]
    assert "Factory Purchase (Import 5%)" in names
    assert "Factory Purchase (Import 18%)" in names
    assert "Tax Free (Purchases)" in names
    assert "Custom Duty Payable" in names
    assert "IGST Purchase @ 5.00 %" in names
    assert "IGST Payable @ 18%" in names


def test_party_ledger_is_title_cased_supplier():
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "9")])
    entries = TallyExporter().build(doc, 95.3)["tallymessage"][0]["allledgerentries"]
    party = next(e for e in entries if e.get("ispartyledger"))
    assert party["ledgername"] == "Prayan Impex Company Limited"


def test_party_amount_equals_total_purchase_inr():
    lines = [make_line(1, 0.05, 100, "KGS", "9"), make_line(2, 0.18, 50, "PCS", "3")]
    doc = make_doc(lines)
    entries = TallyExporter().build(doc, 95.3)["tallymessage"][0]["allledgerentries"]
    party = next(e for e in entries if e.get("ispartyledger"))
    assert abs(float(party["amount"]) - sum(l.purchase_inr for l in lines)) < 0.01


def test_custom_duty_equals_total_duty():
    lines = [make_line(1, 0.05, 100, "KGS", "9"), make_line(2, 0.18, 50, "PCS", "3")]
    doc = make_doc(lines)
    entries = TallyExporter().build(doc, 95.3)["tallymessage"][0]["allledgerentries"]
    duty = next(e for e in entries if e["ledgername"] == "Custom Duty Payable")
    assert abs(float(duty["amount"]) - sum(l.total_customs_duty for l in lines)) < 0.01


def test_output_is_json_serialisable():
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "9")])
    json.dumps(TallyExporter().build(doc, 95.3))


# ---------------------------------------------------------------------------
# Buyer / seller populated FROM THE BOE (the fix in this change)
# ---------------------------------------------------------------------------
def test_buyer_and_seller_come_from_boe_header():
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "9")])
    v = TallyExporter().build(doc, 95.3)["tallymessage"][0]
    # "M/S " prefix stripped for the Tally buyer name.
    assert v["basicbuyername"] == "GEMINI UNICOM LLP"
    assert v["cmpgstin"] == "27AAYFG7003K1ZW"
    assert v["consigneegstin"] == "27AAYFG7003K1ZW"
    assert v["placeofsupply"] == "Maharashtra"
    assert v["consigneepincode"] == "400080"
    assert v["countryofresidence"] == "China"
    assert v["partyname"] == "PRAYAN IMPEX COMPANY LIMITED"
    # Address blocks are populated from the BOE.
    assert v["basicbuyeraddress"][1:] == ["A/43", "DEVIDAYAL ROAD", "MULUND"]
    assert v["address"][1:] == ["ROOM 203", "YIWU"]


def test_profile_overrides_boe_values():
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "9")])
    company = CompanyProfile(name="Custom Buyer Ltd", gstin="24AAAAA0000A1Z5", state="Gujarat")
    seller = SellerProfile(name="Custom Seller Co", country="Vietnam")
    v = TallyExporter(company=company, seller=seller).build(doc, 95.3)["tallymessage"][0]
    assert v["basicbuyername"] == "Custom Buyer Ltd"
    assert v["cmpgstin"] == "24AAAAA0000A1Z5"
    assert v["placeofsupply"] == "Gujarat"
    assert v["partyname"] == "Custom Seller Co"
    assert v["countryofresidence"] == "Vietnam"


def test_missing_buyer_seller_fields_are_blank_not_guessed():
    doc = make_doc(
        [make_line(1, 0.05, 100, "KGS", "9")],
        buyer_gstin=RawValue.missing(),
        buyer_pincode=RawValue.missing(),
        buyer_state=RawValue.missing(),
        seller_country=RawValue.missing(),
    )
    v = TallyExporter().build(doc, 95.3)["tallymessage"][0]
    assert v["cmpgstin"] == ""
    assert v["consigneepincode"] == ""
    assert v["placeofsupply"] == ""


# ---------------------------------------------------------------------------
# Unit conversion (DOZ/GRS/THD -> PCS) on inventory allocations
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("unit,factor", [("DOZ", 12), ("GRS", 144), ("THD", 1000)])
def test_unit_conversion_to_pcs(unit, factor):
    doc = make_doc([make_line(1, 0.05, 10, unit, "9")])
    entries = TallyExporter().build(doc, 95.3)["tallymessage"][0]["allledgerentries"]
    purchase = next(e for e in entries if e["ledgername"] == "Factory Purchase (Import 5%)")
    inv = purchase["inventoryallocations"][0]
    assert inv["actualqty"] == f" {10 * factor:.2f} PCS"
    # Amount is preserved (rate re-derived against the converted qty).
    amount = abs(float(inv["amount"]))
    qty = 10 * factor
    assert abs(float(inv["rate"].split("/")[0]) - amount / qty) < 0.01


def test_non_convertible_unit_is_unchanged():
    doc = make_doc([make_line(1, 0.05, 10, "KGS", "9")])
    entries = TallyExporter().build(doc, 95.3)["tallymessage"][0]["allledgerentries"]
    purchase = next(e for e in entries if e["ledgername"] == "Factory Purchase (Import 5%)")
    inv = purchase["inventoryallocations"][0]
    assert inv["actualqty"] == " 10.00 KGS"


# ---------------------------------------------------------------------------
# apply_stock_names (Step 2 mapping -> in-memory JSON)
# ---------------------------------------------------------------------------
def test_apply_stock_names_overrides_stockitemname():
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "9")])
    mapped = apply_stock_names(doc, {1: "SLIDER (GARMENT ACCESSORY)"})
    entries = TallyExporter().build(mapped, 95.3)["tallymessage"][0]["allledgerentries"]
    purchase = next(e for e in entries if e["ledgername"] == "Factory Purchase (Import 5%)")
    assert purchase["inventoryallocations"][0]["stockitemname"] == "SLIDER (GARMENT ACCESSORY)"


# ---------------------------------------------------------------------------
# Excel round-trip (manual-upload path) + Step 2 name write-back
# ---------------------------------------------------------------------------
def test_excel_reader_reconstructs_lines_without_formulas():
    lines = [make_line(1, 0.05, 100, "KGS", "96071190"), make_line(2, 0.18, 50, "PCS", "39264049")]
    doc = make_doc(lines)
    gen = ExcelGenerator(use_formulas=False)
    xlsx = gen.generate(doc, ReviewFlagSet([]))
    rebuilt = read_workbook(xlsx)
    assert len(rebuilt.lines) == len(lines)
    entries = TallyExporter().build(rebuilt, 95.3)["tallymessage"][0]["allledgerentries"]
    assert abs(_signed_total(entries)) < 0.5


def test_write_tally_names_fills_column_d():
    lines = [make_line(1, 0.05, 100, "KGS", "96071190"), make_line(2, 0.18, 50, "PCS", "39264049")]
    doc = make_doc(lines)
    xlsx = ExcelGenerator(use_formulas=False).generate(doc, ReviewFlagSet([]))
    edited = write_tally_names(xlsx, {1: "SLIDER", 2: "KEYCHAIN"})
    rebuilt = read_workbook(edited)
    # read_workbook prefers column D "AS PER TALLY NAME" for the stock name.
    names = [rebuilt.lines[0].source.description.parsed, rebuilt.lines[1].source.description.parsed]
    assert names == ["SLIDER", "KEYCHAIN"]


def test_excel_reader_rejects_empty_workbook():
    from openpyxl import Workbook

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
    lines = [
        make_line(i + 1, rate, qty, "NOS", "1000")
        for i, (rate, qty) in enumerate(specs)
    ]
    doc = make_doc(lines)
    entries = TallyExporter().build(doc, 95.3)["tallymessage"][0]["allledgerentries"]
    assert abs(_signed_total(entries)) < 0.05
