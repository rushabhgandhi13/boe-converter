"""Tests for the legacy "Indian Customs EDI System V1.5R001" BOE format.

Covers the second supported Bill of Entry layout (dual-format support):

- ``is_old_format`` detection on page text;
- ``UploadValidator`` recognizing the legacy marker set (so old-format uploads
  are not rejected as NOT_A_BOE);
- ``OldFormatParser`` parsing the stacked one-item-per-page line blocks and the
  header fields from synthetic, deterministic page text (no binary asset
  required), including a digit-leading description and a non-DOZ unit (GRS);
- a full end-to-end conversion against the real legacy PDF, skipped when the
  (commercial, un-committed) asset is absent. The legacy document has 122 line
  items, exercising the dynamic overflow positioning (totals + auxiliary tables
  shift below the items so they never overlap).
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest

from boe_converter.old_format_parser import OldFormatParser, is_old_format
from boe_converter.validator import UploadValidator

# ---------------------------------------------------------------------------
# Synthetic legacy-format page text (deterministic; no binary asset needed).
# ---------------------------------------------------------------------------
_OLD_PAGE_HEADER = """Indian Customs EDI System - Imports V1.5R001
V1.5R001
BILL OF ENTRY FOR HOME CONSUMPTION
[Custom Stn: INNSA1]________CHA : AACFI6528FCH001 [INDIAN SHIPPING SERVICES ]
BE No/Dt./cc/Typ:2022823/20/06/2026/N/H
Importer Details :AAOHP8563P PAN : AAOHP8563PFT001 AD Code : 0180792
BL No : OOLU2326556570 H/BL No :
Date : 01/06/2026 Date :
No. Of Pkgs. : 908 CTN Gross Wt. : 23310.000 KGS
Inv No & Dt. : ZJXY26050983 01/06/2026 PRAYAN IMPEX COMPANY LIMITED
Inv Val : 16097.84 USD TOI: CF ROOM 203
Item Details
slno RITC Description RSP Load PROV
Qty Unit Price CTH C.Notn C.NSNO Cus Dty Rt BCD amt(Rs.)
Unit Ass Val CETH E.Notn E.NSNO Exc Dty Rt CVD amt(Rs.)
"""

# Item 1: ordinary description, DOZ unit.
_OLD_ITEM_1 = """1 39264049 KEYCHAIN(O/T REPUTED BRAND)
750.00 0.120000 39264049 15.00 % 1301.00
Cus AIDC 011/2021 17 0.00% 0.00
DOZ 8673.49 NOEXCISE 0.00 % 0.00
Social Welfare Surcharge: 10.00 % 130.10
IGST 009/2025 II127 18.00 % 1818.80
GST Cess 001/2017 56 0.00 % 0.00
Rs. 8673.49 Page Total Rs. 3249.90
"""

# Item 2: digit-leading description ("9v ...") and a non-DOZ unit ("GRS").
_OLD_ITEM_2 = """2 85068090 9v BATTERY(O/T REPUTED BRAND)
1613.50 0.360000 85068090 15.00 % 8396.80
Cus AIDC 011/2021 17 0.00% 0.00
GRS 55978.71 NOEXCISE 0.00 % 0.00
Social Welfare Surcharge: 10.00 % 839.70
IGST 009/2025 II127 18.00 % 11936.90
GST Cess 001/2017 56 0.00 % 0.00
Rs. 55978.71 Page Total Rs. 21173.40
"""


class _FakePage:
    """Minimal stand-in for a pdfplumber page exposing the methods used."""

    def __init__(self, text: str, words: list[dict] | None = None) -> None:
        self._text = text
        self._words = words or []

    def extract_text(self) -> str:
        return self._text

    def extract_words(self):
        return list(self._words)


def _old_pages() -> list[_FakePage]:
    """Two legacy-format pages, one line item each."""
    return [
        _FakePage(_OLD_PAGE_HEADER + _OLD_ITEM_1),
        _FakePage(_OLD_PAGE_HEADER + _OLD_ITEM_2),
    ]


# ---------------------------------------------------------------------------
# Detection
# ---------------------------------------------------------------------------
def test_is_old_format_detects_legacy_markers():
    assert is_old_format(_OLD_PAGE_HEADER) is True


def test_is_old_format_rejects_new_format_text():
    new_format = (
        "BILL OF ENTRY Port Code BE No PART - II PART - III 1.S NO. 2.CTH"
    )
    assert is_old_format(new_format) is False


# ---------------------------------------------------------------------------
# Validator recognizes the legacy format
# ---------------------------------------------------------------------------
class _FakeHandle:
    def __init__(self, pages: list[_FakePage]) -> None:
        self.pages = pages


def test_validator_recognizes_legacy_format():
    validator = UploadValidator()
    handle = _FakeHandle(_old_pages())
    assert validator._is_recognized_boe(handle) is True


def test_validator_rejects_non_boe_text():
    validator = UploadValidator()
    handle = _FakeHandle([_FakePage("just some random invoice text, not a BOE")])
    assert validator._is_recognized_boe(handle) is False


# ---------------------------------------------------------------------------
# OldFormatParser on synthetic pages
# ---------------------------------------------------------------------------
def test_old_parser_extracts_header_fields():
    doc = OldFormatParser().parse(_old_pages(), company_name="Gemini Unicom LLP", usd_rate=95.3)
    h = doc.header
    assert h.be_no.parsed == 2022823 or h.be_no.raw_text == "2022823"
    assert h.be_date.raw_text == "20/06/2026"
    assert h.invoice_no.raw_text == "ZJXY26050983"
    assert h.invoice_date.raw_text == "01/06/2026"
    assert h.invoice_amount.parsed == pytest.approx(16097.84)
    assert h.invoice_currency.raw_text == "USD"
    assert h.package_count.parsed == 908
    assert h.bl_no.raw_text == "OOLU2326556570"
    assert h.bl_date.raw_text == "01/06/2026"


def test_old_parser_extracts_all_line_items_in_order():
    doc = OldFormatParser().parse(_old_pages(), usd_rate=95.3)
    assert len(doc.line_items) == 2
    serials = [li.item_serial for li in doc.line_items]
    assert serials == [1, 2]


def test_old_parser_parses_digit_leading_description_and_non_doz_unit():
    doc = OldFormatParser().parse(_old_pages(), usd_rate=95.3)
    item2 = doc.line_items[1]
    # Digit-leading description must be preserved verbatim.
    assert item2.description.raw_text == "9v BATTERY(O/T REPUTED BRAND)"
    # Non-DOZ unit (GRS) must be captured rather than skipped.
    assert item2.unit.raw_text == "GRS"
    assert item2.quantity.parsed == pytest.approx(1613.50)
    assert item2.assessable_value.parsed == pytest.approx(55978.71)


def test_old_parser_rates_captured_as_fractions():
    """BCD/IGST percentages are stored as decimal fractions (e.g. 18% -> 0.18)."""
    doc = OldFormatParser().parse(_old_pages(), usd_rate=95.3)
    item1 = doc.line_items[0]
    assert item1.igst_rate.parsed == pytest.approx(0.18)
    assert item1.bcd_rate.parsed == pytest.approx(0.15)
    assert item1.bcd_amount.parsed == pytest.approx(1301.00)


def test_old_parser_does_not_flag_absent_total_duty():
    """The legacy format prints no per-line total duty, so it is not flagged."""
    doc = OldFormatParser().parse(_old_pages(), usd_rate=95.3)
    assert not any(f.field_name == "total_duty" for f in doc.flags)


# ---------------------------------------------------------------------------
# Full pipeline against the real legacy PDF (skipped when the asset is absent).
# ---------------------------------------------------------------------------
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
OLD_FORMAT_PDF = _PROJECT_ROOT / "old_format_pdf.pdf"


@pytest.mark.skipif(
    not OLD_FORMAT_PDF.exists(),
    reason=f"legacy-format asset missing: {OLD_FORMAT_PDF.name}",
)
def test_old_format_end_to_end_conversion():
    openpyxl = pytest.importorskip("openpyxl")
    from boe_converter.orchestrator import ConversionOrchestrator

    raw = OLD_FORMAT_PDF.read_bytes()
    orchestrator = ConversionOrchestrator()
    result = orchestrator.convert(raw, "old.pdf", 95.3)

    assert result.ok, f"legacy BOE must convert: {result.error_code} {result.message}"
    summary = result.summary
    assert summary.line_items_extracted == 122
    assert summary.total_invoice_amount_usd == pytest.approx(16097.84, abs=0.01)
    # The declared invoice total matches the summed line amounts (no discrepancy).
    assert summary.review_flag_count == 0

    data = orchestrator.get_download(result.download_token)
    wb = openpyxl.load_workbook(io.BytesIO(data))
    try:
        ws = wb["Sheet1"]
        n = 122
        last_data_row = 13 + n - 1            # 134
        totals_row = 61 + max(0, last_data_row - 60)  # 135
        # Dense item rows 1..122 with no overlap, totals below the last item.
        assert ws.cell(row=13, column=1).value == 1
        assert ws.cell(row=last_data_row, column=1).value == 122
        assert ws.cell(row=totals_row, column=7).value == 908  # total CTN/pkgs
        # The downloaded workbook carries a live SUM formula for the amount total.
        amount_total = ws.cell(row=totals_row, column=12).value
        assert isinstance(amount_total, str) and amount_total.startswith("=SUM(L13:")
    finally:
        wb.close()
