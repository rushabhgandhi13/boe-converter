"""Unit tests for line-item extraction against the reference BOE PDF (task 5.6).

These tests run the real :class:`PdfParser` over the reference ICEGATE Bill of
Entry shipped with the project (``205090022062026INNSA1BE0230620261842.pdf``)
and assert that the Part II (invoice) and Part III (duty) line-item values are
read back at their known printed values. They cover three representative cases:

- **A fully-populated item** (serial 1): every per-line field is asserted
  against its verbatim/parsed value (CTH, description, unit price, quantity,
  UQC, invoice amount, assessable value, BCD rate/amount, SWS amount, IGST rate,
  total duty) (Req 3.4-3.12).
- **The exemption-zero item** (serial 3): a Notification-driven ``0`` BCD
  rate/amount is captured as numeric ``0`` (not blank/missing) (Req 3.14).
- **A wrapped multi-line description item** (serial 45): the description that
  spans more than one printed line is reconstructed verbatim and untruncated
  (Req 3.6).

It also asserts the document-level invariants the merge relies on: exactly 45
line items with a contiguous ``1..45`` serial sequence (Req 3.4).

The reference PDF is large and not always present (e.g. a checkout without the
sample asset), so the whole module is skipped when it is absent.

Validates: Requirements 3.4, 3.5, 3.6, 3.7, 3.8, 3.9, 3.10, 3.11, 3.12, 3.14
"""

from __future__ import annotations

from pathlib import Path

import pytest

pdfplumber = pytest.importorskip("pdfplumber")

from boe_converter.parser import PdfParser  # noqa: E402

# The reference BOE PDF lives at the project root (one level above tests/).
REFERENCE_PDF = (
    Path(__file__).resolve().parent.parent
    / "205090022062026INNSA1BE0230620261842.pdf"
)

# Module-level skip when the reference asset is not available.
pytestmark = pytest.mark.skipif(
    not REFERENCE_PDF.exists(),
    reason=f"reference BOE PDF not found at {REFERENCE_PDF}",
)


@pytest.fixture(scope="module")
def parsed():
    """Parse the reference PDF once and index its line items by serial."""
    parser = PdfParser()
    with pdfplumber.open(str(REFERENCE_PDF)) as pdf:
        doc = parser.parse(pdf, usd_rate=95.3)
    by_serial = {li.item_serial: li for li in doc.line_items}
    return doc, by_serial


@pytest.fixture(scope="module")
def rows():
    """Extract the intermediate Part II/III rows once.

    The invoice ``amount`` and duty ``sws_amount`` fields are carried on the
    intermediate ``InvoiceItemRow``/``DutyItemRow`` (they are inputs to the
    calculator/cross-checks rather than ``LineItem`` fields), so they are read
    here directly from the per-region extractors.
    """
    parser = PdfParser()
    with pdfplumber.open(str(REFERENCE_PDF)) as pdf:
        pages = list(pdf.pages)
        invoice = parser._extract_invoice_items(pages)
        duty = parser._extract_duty_items(pages)
    return invoice, duty


# ---------------------------------------------------------------------------
# Document-level invariants
# ---------------------------------------------------------------------------


def test_extracts_forty_five_items(parsed):
    """Req 3.4 - the reference BOE has exactly 45 line items."""
    doc, _by_serial = parsed
    assert len(doc.line_items) == 45


def test_serials_are_contiguous_one_to_n(parsed):
    """Req 3.4 - item serials form the contiguous sequence 1..45 with no gaps."""
    doc, _by_serial = parsed
    serials = [li.item_serial for li in doc.line_items]
    assert serials == list(range(1, 46))


# ---------------------------------------------------------------------------
# Fully-populated item (serial 1)
# ---------------------------------------------------------------------------


def test_item1_cth_hsn(parsed):
    """Req 3.5 - CTH/HSN is captured verbatim and parsed numerically."""
    _doc, by_serial = parsed
    item = by_serial[1]
    assert item.cth_hsn.raw_text == "42029900"
    assert item.cth_hsn.is_missing is False


def test_item1_description(parsed):
    """Req 3.6 - description is captured verbatim, beginning with the printed
    product name and untruncated."""
    _doc, by_serial = parsed
    item = by_serial[1]
    assert item.description.raw_text == "SHOULDER LADIES BAGS(O/T REPUTED BRAND)"
    assert item.description.raw_text.startswith("SHOULDER LADIES BAGS")
    assert item.description.is_missing is False


def test_item1_unit_price(parsed):
    """Req 3.7 - unit price (UPI) is captured verbatim and parses to 2.4."""
    _doc, by_serial = parsed
    item = by_serial[1]
    assert item.unit_price_usd.raw_text == "2.400000"
    assert item.unit_price_usd.parsed == pytest.approx(2.4)


def test_item1_quantity_and_unit(parsed):
    """Req 3.8 - quantity and UQC are captured verbatim."""
    _doc, by_serial = parsed
    item = by_serial[1]
    assert item.quantity.parsed == pytest.approx(34.0)
    assert item.unit.raw_text == "DOZ"


def test_item1_invoice_amount(rows):
    """Req 3.7/3.8 - the Part II line amount is captured (81.60)."""
    invoice, _duty = rows
    assert invoice[1].amount.parsed == pytest.approx(81.6)


def test_item1_assessable_value(parsed):
    """Req 3.9 - assessable value is captured verbatim and parsed."""
    _doc, by_serial = parsed
    item = by_serial[1]
    assert item.assessable_value.parsed == pytest.approx(7863.97)


def test_item1_bcd_rate_and_amount(parsed):
    """Req 3.10 - BCD rate is captured as a decimal fraction (15% -> 0.15) and
    the BCD amount verbatim."""
    _doc, by_serial = parsed
    item = by_serial[1]
    assert item.bcd_rate.raw_text == "15"
    assert item.bcd_rate.parsed == pytest.approx(0.15)
    assert item.bcd_amount.parsed == pytest.approx(1179.6)


def test_item1_sws_amount(rows):
    """Req 3.10 - the SWS (surcharge) amount is captured (118)."""
    _invoice, duty = rows
    assert duty[1].sws_amount.parsed == pytest.approx(118.0)


def test_item1_igst_rate(parsed):
    """Req 3.11 - IGST rate is captured as a decimal fraction (18% -> 0.18)."""
    _doc, by_serial = parsed
    item = by_serial[1]
    assert item.igst_rate.raw_text == "18"
    assert item.igst_rate.parsed == pytest.approx(0.18)


def test_item1_total_duty(parsed):
    """Req 3.12 - total duty is captured verbatim and parsed."""
    _doc, by_serial = parsed
    item = by_serial[1]
    assert item.total_duty.parsed == pytest.approx(2946.7)


# ---------------------------------------------------------------------------
# Exemption-zero item (serial 3)
# ---------------------------------------------------------------------------


def test_item3_exemption_zero_bcd_is_numeric_zero(parsed):
    """Req 3.14 - an exemption-driven 0 BCD rate/amount is captured as numeric
    0, never blank or missing."""
    _doc, by_serial = parsed
    item = by_serial[3]

    assert item.bcd_rate.is_missing is False
    assert item.bcd_rate.is_unparseable is False
    assert item.bcd_rate.parsed == 0
    assert item.bcd_rate.parsed == pytest.approx(0.0)

    assert item.bcd_amount.is_missing is False
    assert item.bcd_amount.is_unparseable is False
    assert item.bcd_amount.parsed == 0
    assert item.bcd_amount.parsed == pytest.approx(0.0)


def test_item3_sws_amount_zero(rows):
    """Req 3.14 - the dependent SWS amount is likewise captured as numeric 0."""
    _invoice, duty = rows
    assert duty[3].sws_amount.parsed == pytest.approx(0.0)


def test_item3_other_fields_still_populated(parsed):
    """The exemption case zeroes only BCD; other fields remain populated."""
    _doc, by_serial = parsed
    item = by_serial[3]
    assert item.assessable_value.parsed == pytest.approx(104081.9)
    assert item.igst_rate.parsed == pytest.approx(0.18)
    assert item.total_duty.parsed == pytest.approx(18734.7)


# ---------------------------------------------------------------------------
# Wrapped multi-line description item (serial 45)
# ---------------------------------------------------------------------------


def test_item45_wrapped_description_reconstructed_verbatim(parsed):
    """Req 3.6 - the multi-line (wrapped) description is stitched back together
    verbatim and untruncated, with each fragment appearing once."""
    _doc, by_serial = parsed
    item = by_serial[45]
    assert (
        item.description.raw_text
        == "EXAMINATION SET OF 4 (DENTAL KIT ACCESSORY)(O/T REPUTED BRAND)"
    )
    assert item.description.is_missing is False


def test_item45_remaining_fields(parsed):
    """Req 3.5-3.12 - the wrapped item's other fields are captured correctly."""
    _doc, by_serial = parsed
    item = by_serial[45]
    assert item.cth_hsn.raw_text == "90184900"
    assert item.unit_price_usd.parsed == pytest.approx(0.2)
    assert item.quantity.parsed == pytest.approx(81.0)
    assert item.unit.raw_text == "PCS"
    assert item.assessable_value.parsed == pytest.approx(1561.23)
    assert item.bcd_rate.parsed == pytest.approx(0.075)
    assert item.bcd_amount.parsed == pytest.approx(117.1)
    assert item.igst_rate.parsed == pytest.approx(0.05)
    assert item.total_duty.parsed == pytest.approx(303.5)
