"""Unit tests for header-block extraction against the reference BOE PDF (task 4.5).

These tests run the real :meth:`PdfParser._extract_header` over the reference
ICEGATE Bill of Entry shipped with the project
(``205090022062026INNSA1BE0230620261842.pdf``) and assert every Req 2.1-2.9
header field is read back at its known printed value. They also assert that an
empty document (no pages) produces a ``MISSING`` :class:`ReviewFlag` for each
required field rather than substituting a default (Req 2.10).

The reference PDF is large and not always present (e.g. a checkout without the
sample asset), so the whole module is skipped when it is absent.

Validates: Requirements 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.8, 2.9, 2.10
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
def header():
    """Extract the Header_Block from the reference PDF exactly once."""
    parser = PdfParser()
    with pdfplumber.open(str(REFERENCE_PDF)) as pdf:
        block, flags = parser._extract_header(
            list(pdf.pages), company_name="Gemini Unicom LLP", usd_rate=95.3
        )
    return block, flags


def test_no_review_flags_for_reference_pdf(header):
    """Every required Req 2.x field is locatable in the reference PDF, so the
    header extraction raises no review flags (Req 2.10 - nothing missing)."""
    _block, flags = header
    assert flags == [], f"unexpected header review flags: {flags}"


def test_be_no_matches_reference(header):
    """Req 2.1 - BE No is read verbatim and parses to its numeric value."""
    block, _flags = header
    assert block.be_no.raw_text == "2050900"
    assert block.be_no.parsed == 2050900.0
    assert block.be_no.is_missing is False


def test_be_date_matches_reference(header):
    """Req 2.2 - BE Date is captured verbatim as printed."""
    block, _flags = header
    assert block.be_date.raw_text == "22/06/2026"
    assert block.be_date.is_missing is False


def test_invoice_no_matches_reference(header):
    """Req 2.3 - Invoice No is read verbatim."""
    block, _flags = header
    assert block.invoice_no.raw_text == "202605105"
    assert block.invoice_no.is_missing is False


def test_invoice_date_matches_reference(header):
    """Req 2.4 - Invoice Date is captured verbatim as printed (Part II form)."""
    block, _flags = header
    assert block.invoice_date.raw_text == "07-JUN-26"
    assert block.invoice_date.is_missing is False


def test_invoice_amount_matches_reference(header):
    """Req 2.5 - invoice amount is captured verbatim and parsed numerically."""
    block, _flags = header
    assert block.invoice_amount.raw_text == "20671.32"
    assert block.invoice_amount.parsed == pytest.approx(20671.32)
    assert block.invoice_amount.is_missing is False


def test_invoice_currency_matches_reference(header):
    """Req 2.5 - invoice currency is captured verbatim."""
    block, _flags = header
    assert block.invoice_currency.raw_text == "USD"
    assert block.invoice_currency.is_missing is False


def test_package_count_matches_reference(header):
    """Req 2.6 - total package/CTN count is captured and parses to its number."""
    block, _flags = header
    assert block.package_count.raw_text == "1357"
    assert block.package_count.parsed == 1357.0
    assert block.package_count.is_missing is False


def test_party_name_matches_reference(header):
    """Req 2.7 - supplier party name is captured from the supplier column only."""
    block, _flags = header
    assert block.party_name.raw_text == "PRAYAN IMPEX COMPANY LIMITED"
    assert block.party_name.is_missing is False


def test_container_details_match_reference(header):
    """Req 2.8 - container details include the printed container number."""
    block, _flags = header
    assert block.container_details.is_missing is False
    assert "NBYU8258966" in (block.container_details.raw_text or "")


def test_bl_no_and_date_match_reference(header):
    """Req 2.9 - Bill of Lading No/Date are extracted where present."""
    block, _flags = header
    assert block.bl_no.raw_text == "NOSNB26NS1"
    assert block.bl_no.is_missing is False
    assert block.bl_date.raw_text == "07/06/2026"
    assert block.bl_date.is_missing is False


# Required header fields per Req 2.1-2.8 (B/L 2.9 is "where present" and is not
# flagged when absent, so it is intentionally excluded here).
REQUIRED_HEADER_FIELDS = [
    "be_no",
    "be_date",
    "invoice_no",
    "invoice_date",
    "invoice_amount",
    "invoice_currency",
    "package_count",
    "party_name",
    "container_details",
]


def test_empty_pages_flag_every_required_field_missing():
    """Req 2.10 - when no field can be located (empty document), each required
    field yields a MISSING ReviewFlag and no default value is substituted."""
    parser = PdfParser()
    block, flags = parser._extract_header(
        [], company_name="Gemini Unicom LLP", usd_rate=95.3
    )

    missing_fields = {
        f.field_name for f in flags if f.scope == "header" and f.reason == "MISSING"
    }
    for name in REQUIRED_HEADER_FIELDS:
        assert name in missing_fields, f"expected MISSING flag for {name}"
        assert getattr(block, name).is_missing is True

    # B/L (Req 2.9) is "where present" - it must never be flagged when absent.
    bl_flagged = {
        f.field_name for f in flags if f.field_name in ("bl_no", "bl_date")
    }
    assert bl_flagged == set()
