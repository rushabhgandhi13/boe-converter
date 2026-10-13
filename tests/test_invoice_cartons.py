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

from boe_converter.excel_writer import (
    COL_CTN,
    COL_RATE_PER_UNIT,
    COL_SR_NO,
    ExcelGenerator,
    _totals_row_for,
)
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
    def __init__(self, cartons=None, descriptions=None, unit_prices=None,
                 quantities=None, amounts=None):
        self._cartons = dict(cartons or {})
        self._descriptions = dict(descriptions or {})
        self._unit_prices = dict(unit_prices or {})
        self._quantities = dict(quantities or {})
        self._amounts = dict(amounts or {})

    def parse_line_details(self, doc):  # noqa: ANN001
        serials = (
            set(self._cartons) | set(self._descriptions) | set(self._unit_prices)
            | set(self._quantities) | set(self._amounts)
        )
        return {
            s: {
                "cartons": self._cartons.get(s),
                "description": self._descriptions.get(s),
                "unit_price": self._unit_prices.get(s),
                "quantity": self._quantities.get(s),
                "amount": self._amounts.get(s),
            }
            for s in serials
        }


class _RaisingInvoiceParser:
    def parse_line_details(self, doc):  # noqa: ANN001
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


def test_attach_invoice_overrides_description():
    """An invoice description replaces the BOE description (even without cartons)."""
    orch = ConversionOrchestrator(
        invoice_parser=_StubInvoiceParser(
            cartons={1: _num(5)},
            descriptions={1: _num("INVOICE NAME ONE"), 2: _num("INVOICE NAME TWO")},
        )
    )
    enriched = orch._attach_cartons(_extracted(2), b"%PDF-fake")
    by_serial = {li.item_serial: li for li in enriched.line_items}
    assert by_serial[1].description.parsed == "INVOICE NAME ONE"
    assert by_serial[1].cartons.parsed == 5
    # Line 2 gets the invoice description even though it has no carton count.
    assert by_serial[2].description.parsed == "INVOICE NAME TWO"
    assert by_serial[2].cartons.is_missing is True


def test_attach_invoice_keeps_boe_description_when_invoice_has_none():
    """A line absent from the invoice keeps its BOE description."""
    orch = ConversionOrchestrator(
        invoice_parser=_StubInvoiceParser(descriptions={1: _num("ONLY ONE")})
    )
    enriched = orch._attach_cartons(_extracted(2), b"%PDF-fake")
    by_serial = {li.item_serial: li for li in enriched.line_items}
    assert by_serial[1].description.parsed == "ONLY ONE"
    assert by_serial[2].description.parsed == "item 2"  # unchanged BOE description


def test_attach_invoice_keeps_boe_unit_price_when_invoice_has_price():
    """Unit Price in USD is sourced from the BOE, never the optional invoice."""
    orch = ConversionOrchestrator(
        invoice_parser=_StubInvoiceParser(unit_prices={1: _num(0.12)})
    )
    enriched = orch._attach_cartons(_extracted(1), b"%PDF-fake")
    assert enriched.line_items[0].unit_price_usd.parsed == 2.0


def _mline(serial: int, qty: float, price: float) -> LineItem:
    """A line with controllable quantity / unit price for mismatch tests."""
    from dataclasses import replace

    return replace(_line(serial), quantity=_num(qty), unit_price_usd=_num(price))


def _mextracted() -> ExtractedDocument:
    return ExtractedDocument(
        header=_header(),
        line_items=[_mline(1, 10.0, 2.0), _mline(2, 7.5, 1.2), _mline(3, 4.0, 0.1)],
        declared_item_count=3,
    )


def test_matching_values_leave_no_mismatch():
    """Identical BOE/invoice numbers are not flagged."""
    orch = ConversionOrchestrator(
        invoice_parser=_StubInvoiceParser(
            quantities={1: _num(10.0)},
            unit_prices={1: _num(2.0)},
            amounts={1: _num(20.0)},
        )
    )
    enriched = orch._attach_cartons(_mextracted(), b"%PDF-fake")
    by_serial = {li.item_serial: li for li in enriched.line_items}
    assert by_serial[1].invoice_mismatch == ()
    assert by_serial[2].invoice_mismatch == ()  # absent from invoice
    assert by_serial[3].invoice_mismatch == ()


def test_differing_price_and_amount_are_flagged_per_field():
    """Only the fields that actually differ are named (amount follows)."""
    orch = ConversionOrchestrator(
        invoice_parser=_StubInvoiceParser(
            quantities={2: _num(7.5)},
            unit_prices={2: _num(9.99)},
            amounts={2: _num(9.0)},
        )
    )
    enriched = orch._attach_cartons(_mextracted(), b"%PDF-fake")
    by_serial = {li.item_serial: li for li in enriched.line_items}
    # BOE 1.2 x 7.5 = 9.0 agrees (within float dust); only the price differs.
    assert by_serial[2].invoice_mismatch == ("unit_price_usd",)


def test_float_dust_in_recomputed_amount_is_not_flagged():
    """0.1 x 4.0 recomputes with binary dust; it must not turn red."""
    orch = ConversionOrchestrator(
        invoice_parser=_StubInvoiceParser(
            quantities={3: _num(4.0)},
            unit_prices={3: _num(0.1)},
            amounts={3: _num(0.4)},
        )
    )
    enriched = orch._attach_cartons(_mextracted(), b"%PDF-fake")
    by_serial = {li.item_serial: li for li in enriched.line_items}
    assert by_serial[3].invoice_mismatch == ()


def test_missing_sides_skip_comparison():
    """A field with no number on either side is skipped, never flagged."""
    orch = ConversionOrchestrator(
        invoice_parser=_StubInvoiceParser(quantities={1: _num(10.0)})
    )
    enriched = orch._attach_cartons(_mextracted(), b"%PDF-fake")
    by_serial = {li.item_serial: li for li in enriched.line_items}
    assert by_serial[1].invoice_mismatch == ()


def test_half_cent_rounding_in_amount_is_not_flagged():
    """8333.5 x 0.15 = 1250.025 prints as 1250.03; that is rounding, not a
    mismatch (live BE 4270364 rows 8/15 false alarm)."""
    from dataclasses import replace

    base = _mline(1, 8333.5, 0.15)
    doc = ExtractedDocument(
        header=_header(), line_items=[base], declared_item_count=1
    )
    orch = ConversionOrchestrator(
        invoice_parser=_StubInvoiceParser(
            quantities={1: _num(8333.5)},
            unit_prices={1: _num(0.15)},
            amounts={1: _num(1250.03)},
        )
    )
    enriched = orch._attach_cartons(doc, b"%PDF-fake")
    assert enriched.line_items[0].invoice_mismatch == ()


def test_genuine_one_cent_amount_difference_is_flagged():
    """Same qty/price but the amount off by a full cent is a real mismatch."""
    from dataclasses import replace

    base = _mline(1, 10.0, 2.0)  # recomputed 20.00
    doc = ExtractedDocument(
        header=_header(), line_items=[base], declared_item_count=1
    )
    orch = ConversionOrchestrator(
        invoice_parser=_StubInvoiceParser(
            quantities={1: _num(10.0)},
            unit_prices={1: _num(2.0)},
            amounts={1: _num(20.01)},
        )
    )
    enriched = orch._attach_cartons(doc, b"%PDF-fake")
    assert enriched.line_items[0].invoice_mismatch == ("amount_usd",)


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


def _doc_with_cartons(cartons, package=908):
    """A one-to-N-line doc; ``cartons`` is [(serial, count|None)]."""
    from dataclasses import replace

    lines = []
    for serial, ctn in cartons:
        item = _line(serial)
        if ctn is not None:
            item = replace(item, cartons=_num(ctn))
        lines.append(ComputedLine(source=item))
    return ComputedDocument(
        header=_header(),
        lines=lines,
        totals=Totals(package_count=_num(package)),
    )


def test_ctn_total_is_sum_formula_when_cartons_present():
    ws = ExcelGenerator(use_formulas=True).build_workbook(
        _doc_with_cartons([(1, 10), (2, 20), (3, 30)]), ReviewFlagSet()
    )["Sheet1"]
    # 3 items -> rows 13..15, two blanks, totals at row 18.
    assert ws.cell(row=_totals_row_for(3), column=COL_CTN).value == "=SUM(G13:G15)"


def test_ctn_total_is_numeric_sum_in_literal_mode():
    ws = ExcelGenerator(use_formulas=False).build_workbook(
        _doc_with_cartons([(1, 10), (2, 20), (3, 30)]), ReviewFlagSet()
    )["Sheet1"]
    assert ws.cell(row=_totals_row_for(3), column=COL_CTN).value == 60


def test_ctn_total_falls_back_to_package_count_without_cartons():
    ws = ExcelGenerator(use_formulas=False).build_workbook(
        _doc_with_cartons([(1, None)]), ReviewFlagSet()
    )["Sheet1"]
    assert ws.cell(row=_totals_row_for(1), column=COL_CTN).value == 908


def _is_mismatch_red(ws, row: int, col: int) -> bool:
    fill = ws.cell(row=row, column=col).fill
    return (
        fill is not None
        and getattr(fill, "fgColor", None) is not None
        and str(fill.fgColor.rgb) == "FFFF0000"
    )


def test_mismatched_row_is_red_filled_clean_row_is_not():
    from dataclasses import replace

    flagged = replace(_line(2), invoice_mismatch=("unit_price_usd",))
    doc = ComputedDocument(
        header=_header(),
        lines=[ComputedLine(source=_line(1)), ComputedLine(source=flagged)],
        totals=Totals(package_count=_num(908)),
    )
    ws = ExcelGenerator(use_formulas=False).build_workbook(doc, ReviewFlagSet())["Sheet1"]
    cols = range(COL_SR_NO, COL_RATE_PER_UNIT + 1)
    assert all(not _is_mismatch_red(ws, 13, c) for c in cols)
    assert all(_is_mismatch_red(ws, 14, c) for c in cols)


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


@pytest.mark.skipif(not INVOICE_PDF.exists(), reason="invoice asset missing")
def test_invoice_parser_reads_descriptions():
    details = InvoicePackingListParser().parse_line_details(str(INVOICE_PDF))
    assert details[1]["description"].parsed == "SLIDERS (GARMENT ACCESSORY)"
    assert details[2]["description"].parsed == "KEYCHAIN"
    assert details[3]["description"].parsed == "NAIL CUTTER"


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
