"""Tests for three field-extraction fixes on the new ICEGATE BOE format.

1. Company name (Excel E1) is read from the BOE's importer name (Part I Section
   B "1.IMPORTER NAME & ADDRESS"), not left at the configured default.
2. The Master/House B/L number (MAWB/HAWB) can wrap onto a second line in the
   PDF (e.g. "NGBCB26011" + "430"); the full value must be reassembled.
3. An additional customs cess (Part III "2.CHCESS") - present on only some BOEs
   - is folded into the customs-duty base: CUST AIDC (Excel col U) = BCD amount
   + CHCESS, with SWS = 10% of that combined base.

Calculator behaviour (fix 3) is unit-tested deterministically; the extraction
fixes are exercised end-to-end against the real BOE when the (commercial,
un-committed) asset is present.
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest

from boe_converter.calculator import ValueCalculator
from boe_converter.models import LineItem, RawValue


def _num(v) -> RawValue:
    return RawValue(raw_text=str(v), parsed=v)


def _line(**over) -> LineItem:
    base = dict(
        item_serial=1,
        cth_hsn=_num("90191090"),
        description=RawValue(raw_text="X", parsed="X"),
        unit_price_usd=_num(2.5),
        quantity=_num(262.5),
        unit=RawValue(raw_text="DOZ", parsed="DOZ"),
        assessable_value=_num(62514.21),
        bcd_rate=_num(0.075),
        bcd_amount=_num(4688.6),
        igst_rate=_num(0.05),
        total_duty=_num(12151.2),
    )
    base.update(over)
    return LineItem(**base)


# ---------------------------------------------------------------------------
# Fix 3: CHCESS folded into the customs-duty base (deterministic)
# ---------------------------------------------------------------------------
def test_chcess_added_to_cust_aidc_and_sws_base():
    """CUST AIDC = BCD + CHCESS; SWS = 10% of that; matches the sample workbook."""
    item = _line(chcess_amount=_num(3125.71))
    line = ValueCalculator().compute_line(item, 94.2)

    assert line.cust_aidc == pytest.approx(7814.31)          # 4688.6 + 3125.71
    assert line.sws_amount == pytest.approx(781.431)         # 10% of cust_aidc
    assert line.total_customs_duty == pytest.approx(8595.741)  # cust_aidc + sws


def test_no_chcess_leaves_cust_aidc_equal_to_bcd():
    """Without CHCESS the base is just the BCD amount (no regression)."""
    line = ValueCalculator().compute_line(_line(), 94.2)
    assert line.cust_aidc == pytest.approx(4688.6)
    assert line.sws_amount == pytest.approx(468.86)
    assert line.total_customs_duty == pytest.approx(5157.46)


def test_chcess_missing_input_treated_as_zero():
    """A missing/blank CHCESS contributes nothing."""
    item = _line(chcess_amount=RawValue.missing())
    line = ValueCalculator().compute_line(item, 94.2)
    assert line.cust_aidc == pytest.approx(4688.6)


# ---------------------------------------------------------------------------
# End-to-end extraction fixes (real BOE; skipped when the asset is absent)
# ---------------------------------------------------------------------------
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
BE_PDF = _PROJECT_ROOT / "PRATIK C SANGHAVI (HUF) 1071 CTN BE.pdf"

pytestmark_assets = pytest.mark.skipif(
    not BE_PDF.exists(), reason=f"BOE asset missing: {BE_PDF.name}"
)


@pytestmark_assets
def test_company_name_extracted_from_importer():
    from boe_converter.parser import PdfParser

    doc = PdfParser().parse(io.BytesIO(BE_PDF.read_bytes()), usd_rate=94.2)
    assert doc.header.company_name == "M/S PRATIK C SANGHAVI (HUF)"


@pytestmark_assets
def test_bl_number_not_truncated_when_wrapped():
    from boe_converter.parser import PdfParser

    doc = PdfParser().parse(io.BytesIO(BE_PDF.read_bytes()), usd_rate=94.2)
    # The MAWB wraps as "NGBCB26011" + "430"; the full value must be present.
    assert doc.header.bl_no.raw_text == "NGBCB26011430"


@pytestmark_assets
def test_chcess_extracted_for_item_with_additional_cess():
    from boe_converter.parser import PdfParser

    doc = PdfParser().parse(io.BytesIO(BE_PDF.read_bytes()), usd_rate=94.2)
    by_serial = {li.item_serial: li for li in doc.line_items}
    # Item 12 (HANDHELD BODY MASSAGER) carries CHCESS = 3125.71.
    assert by_serial[12].chcess_amount.parsed == pytest.approx(3125.71)
    # An ordinary item has no CHCESS.
    assert by_serial[1].chcess_amount.is_missing
