"""Tests for optional per-line carton (CTN) enrichment from the supplier invoice.

The Bill of Entry has no per-line carton count; when the supplier invoice (an
"invoice cum packing list" PDF) is supplied, its ``TOTAL CTNS`` column is read
and matched to the BOE line items by serial so Excel column ``G`` is populated.
When no invoice is supplied (or it fails to parse) the CTN column stays blank.

Deterministic unit tests use synthetic models / stubs; an integration test runs
the real invoice PDF when the (commercial, un-committed) asset is present.
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest

from boe_converter.excel_writer import COL_CTN, ExcelGenerator
from boe_converter.invoice_parser import InvoicePackingListParser
from boe_converter.models import (
    ComputedDocument,
    ComputedLine,
    ExtractedDocument,
    HeaderBlock,
    LineItem,
    RawValue,
    ReviewFlagSet,
    Totals,
)
from boe_converter.orchestrator import ConversionOrchestrator


def _num(v) -> RawValue:
    return RawValue(raw_text=str(v), parsed=v)


def _line(serial: int) -> LineItem:
    return LineItem(
        item_serial=serial,
        cth_hsn=_num(12345678),
        description=RawValue(raw_text=f"item {serial}", parsed=f"item {serial}"),
        unit_price_usd=_num(2.0),
        quantity=_num(10.0),
        unit=RawValue(raw_text="PCS", parsed="PCS"),
        assessable_value=_num(1000.0),
        bcd_rate=_num(0.10),
        bcd_amount=_num(100.0),
        igst_rate=_num(0.18),
        total_duty=_num(50.0),
    )


def _header() -> HeaderBlock:
    return HeaderBlock(
        company_name="Gemini Unicom LLP", party_name=_num("X"), usd_rate=95.3,
        details=RawValue.missing(), invoice_no=_num("INV"), invoice_date=_num("01/06/2026"),
        be_no=_num("BE"), be_date=_num("01/06/2026"), bl_no=_num("BL"),
        bl_date=_num("01/06/2026"), invoice_amount=_num(100.0), invoice_currency=_num("USD"),
        package_count=_num(908), container_details=_num("C1"),
    )


def _extracted(n: int) -> ExtractedDocument:
    return ExtractedDocument(
        header=_header(),
        line_items=[_line(s) for s in range(1, n + 1)],
        declared_item_count=n,
    )


# ---------------------------------------------------------------------------
# LineItem model
# ---------------------------------------------------------------------------
def test_lineitem_cartons_defaults_to_missing():
    assert _line(1).cartons.is_missing is True


# ---------------------------------------------------------------------------
# Orchestrator carton enrichment (stubbed invoice parser, no real PDF)
# ---------------------------------------------------------------------------
class _StubInvoiceParser:
    def __init__(self, mapping):
        self._mapping = mapping

    def parse_cartons(self, doc):  # noqa: ANN001
        return dict(self._mapping)


class _RaisingInvoiceParser:
    def parse_cartons(self, doc):  # noqa: ANN001
        raise ValueError("bad invoice")


def test_attach_cartons_sets_only_matching_serials():
    orch = ConversionOrchestrator(
        invoice_parser=_StubInvoiceParser({1: _num(131), 3: _num(11)})
    )
    enriched = orch._attach_cartons(_extracted(3), b"%PDF-fake")
    by_serial = {li.item_serial: li for li in enriched.line_items}
    assert by_serial[1].cartons.parsed == 131
    assert by_serial[2].cartons.is_missing is True  # absent from invoice
    assert by_serial[3].cartons.parsed == 11


def test_attach_cartons_is_non_fatal_on_parse_error():
    orch = ConversionOrchestrator(invoice_parser=_RaisingInvoiceParser())
    extracted = _extracted(2)
    enriched = orch._attach_cartons(extracted, b"%PDF-bad")
    # Unchanged: every line keeps a blank carton cell, no exception raised.
    assert all(li.cartons.is_missing for li in enriched.line_items)


def test_attach_cartons_empty_mapping_leaves_lines_unchanged():
    orch = ConversionOrchestrator(invoice_parser=_StubInvoiceParser({}))
    enriched = orch._attach_cartons(_extracted(2), b"%PDF-fake")
    assert all(li.cartons.is_missing for li in enriched.line_items)


# ---------------------------------------------------------------------------
# Excel writer: per-line CTN cell (column G)
# ---------------------------------------------------------------------------
def test_excel_writes_per_line_cartons_when_present():
    from dataclasses import replace

    item = replace(_line(1), cartons=_num(131))
    doc = ComputedDocument(
        header=_header(),
        lines=[ComputedLine(source=item)],
        totals=Totals(package_count=_num(908)),
    )
    ws = ExcelGenerator(use_formulas=False).build_workbook(doc, ReviewFlagSet())["Sheet1"]
    assert ws.cell(row=13, column=COL_CTN).value == 131


def test_excel_leaves_cartons_blank_when_missing():
    doc = ComputedDocument(
        header=_header(),
        lines=[ComputedLine(source=_line(1))],
        totals=Totals(package_count=_num(908)),
    )
    ws = ExcelGenerator(use_formulas=False).build_workbook(doc, ReviewFlagSet())["Sheet1"]
    assert ws.cell(row=13, column=COL_CTN).value is None


# ---------------------------------------------------------------------------
# Real invoice PDF integration (skipped when the asset is absent)
# ---------------------------------------------------------------------------
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
INVOICE_PDF = _PROJECT_ROOT / "INV&PKG_LIST.pdf"
OLD_FORMAT_PDF = _PROJECT_ROOT / "old_format_pdf.pdf"


@pytest.mark.skipif(not INVOICE_PDF.exists(), reason="invoice asset missing")
def test_invoice_parser_reads_total_ctns_column():
    cartons = InvoicePackingListParser().parse_cartons(str(INVOICE_PDF))
    # Spot checks against the invoice's TOTAL CTNS column.
    assert cartons[1].parsed == 131
    assert cartons[2].parsed == 15
    assert cartons[122].parsed == 1
    # Blank carton cells in the invoice produce no entry.
    assert 37 not in cartons
    assert 52 not in cartons


@pytest.mark.skipif(
    not (INVOICE_PDF.exists() and OLD_FORMAT_PDF.exists()),
    reason="invoice and/or BOE asset missing",
)
def test_conversion_with_invoice_populates_ctn_column():
    openpyxl = pytest.importorskip("openpyxl")
    orch = ConversionOrchestrator()
    boe = OLD_FORMAT_PDF.read_bytes()
    inv = INVOICE_PDF.read_bytes()

    result = orch.convert(boe, "boe.pdf", 95.3, invoice_raw=inv)
    assert result.ok
    wb = openpyxl.load_workbook(io.BytesIO(orch.get_download(result.download_token)))
    ws = wb["Sheet1"]
    # First three BOE lines pick up the invoice cartons (matched by serial).
    assert ws.cell(row=13, column=COL_CTN).value == 131
    assert ws.cell(row=14, column=COL_CTN).value == 15
    assert ws.cell(row=15, column=COL_CTN).value == 11


@pytest.mark.skipif(not OLD_FORMAT_PDF.exists(), reason="BOE asset missing")
def test_conversion_without_invoice_leaves_ctn_blank():
    openpyxl = pytest.importorskip("openpyxl")
    orch = ConversionOrchestrator()
    boe = OLD_FORMAT_PDF.read_bytes()

    result = orch.convert(boe, "boe.pdf", 95.3)
    assert result.ok
    wb = openpyxl.load_workbook(io.BytesIO(orch.get_download(result.download_token)))
    ws = wb["Sheet1"]
    assert ws.cell(row=13, column=COL_CTN).value is None
