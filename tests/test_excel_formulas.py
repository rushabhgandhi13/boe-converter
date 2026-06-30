"""Tests for live Excel formulas in the generated workbook (formula mode).

The sample workbook ``1357 ctn llp.xlsx`` stores its computed columns as live
Excel formulas (e.g. ``L13=H13*K13``, ``M13=L13*95.3``, ``Y13=O13/H13``) and its
Totals_Row as ``=SUM(...)`` over the data rows, while the directly-extracted /
manually-keyed columns (``N`` assessable value, ``Q`` IGST, ``U`` BCD amount,
``V`` SWS rate, ``H`` qty, ``K`` unit price) are literals. ``ExcelGenerator``
reproduces this in its default (formula) mode so the downloaded workbook
recalculates in Excel exactly like the sample.

These tests assert:

- the per-line computed cells carry the sample's exact formula strings;
- the ``pcs`` (J) formula appears only on ``DOZ`` rows;
- the literal columns remain plain values (no formula);
- the Totals_Row carries ``=SUM(<col>13:<col>N)`` over the actual data rows and
  the ``total usd`` aux cell mirrors ``=L<totals_row>``;
- a missing/non-numeric input leaves the computed cell blank (no formula that
  would read a blank input as 0);
- literal mode (``use_formulas=False``) emits the evaluated numbers instead, and
  those numbers match the formulas' arithmetic.
"""

from __future__ import annotations

from boe_converter.excel_writer import (
    COL_AMOUNT,
    COL_GST,
    COL_LAND_COST_WITH_GST,
    COL_LAND_COST_WITHOUT_GST,
    COL_PCS,
    COL_RATE_OF_DUTY_IGST,
    COL_RATE_OF_INTEREST_BCD,
    COL_RATE_PER_UNIT,
    COL_RATE_PER_USD,
    COL_SURCHARGE,
    COL_TOTAL_CUSTOM_DUTY,
    COL_TOTAL_CUSTOM_DUTY_2,
    ExcelGenerator,
)
from boe_converter.calculator import ValueCalculator
from boe_converter.models import (
    ExtractedDocument,
    HeaderBlock,
    LineItem,
    RawValue,
    ReviewFlagSet,
)


def _num(value) -> RawValue:
    return RawValue(raw_text=str(value), parsed=float(value))


def _str(value: str) -> RawValue:
    return RawValue(raw_text=value, parsed=value)


def _header(usd_rate: float = 95.3) -> HeaderBlock:
    return HeaderBlock(
        company_name="Gemini Unicom LLP",
        party_name=_str("ACME"),
        usd_rate=usd_rate,
        details=RawValue.missing(),
        invoice_no=_str("INV1"),
        invoice_date=_str("01/06/2026"),
        be_no=_str("BE1"),
        be_date=_str("01/06/2026"),
        bl_no=_str("BL1"),
        bl_date=_str("01/06/2026"),
        invoice_amount=_num(100),
        invoice_currency=_str("USD"),
        package_count=_num(500),
        container_details=_str("C1"),
    )


def _item(
    serial: int,
    *,
    unit: str = "DOZ",
    qty=10,
    price=2,
    missing_price=False,
    bcd_rate=0.10,
    igst_rate=0.18,
) -> LineItem:
    return LineItem(
        item_serial=serial,
        cth_hsn=_str("12345678"),
        description=_str(f"item {serial}"),
        unit_price_usd=RawValue.missing() if missing_price else _num(price),
        quantity=_num(qty),
        unit=_str(unit),
        assessable_value=_num(1000),
        bcd_rate=_num(bcd_rate),
        bcd_amount=_num(100),
        igst_rate=_num(igst_rate),
        total_duty=_num(50),
    )


def _build(items, usd_rate: float = 95.3, use_formulas: bool = True):
    doc = ExtractedDocument(header=_header(usd_rate), line_items=items)
    computed = ValueCalculator().compute(doc, usd_rate)
    wb = ExcelGenerator(use_formulas=use_formulas).build_workbook(computed, ReviewFlagSet())
    return wb["Sheet1"], computed


def test_per_line_formulas_match_sample_pattern():
    ws, _ = _build([_item(1, unit="DOZ")])
    r = 13
    assert ws.cell(row=r, column=COL_PCS).value == f"=H{r}*12"
    assert ws.cell(row=r, column=COL_AMOUNT).value == f"=H{r}*K{r}"
    assert ws.cell(row=r, column=COL_RATE_PER_USD).value == f"=L{r}*95.3"
    assert ws.cell(row=r, column=COL_SURCHARGE).value == f"=U{r}*V{r}"
    assert ws.cell(row=r, column=COL_TOTAL_CUSTOM_DUTY).value == f"=U{r}+W{r}"
    assert ws.cell(row=r, column=COL_LAND_COST_WITHOUT_GST).value == f"=M{r}+P{r}"
    assert ws.cell(row=r, column=COL_TOTAL_CUSTOM_DUTY_2).value == f"=P{r}+Q{r}"
    assert ws.cell(row=r, column=COL_LAND_COST_WITH_GST).value == f"=M{r}+S{r}"
    assert ws.cell(row=r, column=COL_RATE_PER_UNIT).value == f"=O{r}/H{r}"


def test_pcs_formula_only_on_doz_rows():
    ws, _ = _build([_item(1, unit="DOZ"), _item(2, unit="PCS")])
    assert ws.cell(row=13, column=COL_PCS).value == "=H13*12"
    assert ws.cell(row=14, column=COL_PCS).value is None


def test_igst_column_q_is_literal_not_formula():
    """Column Q (GST/IGST) is a manually-keyed literal in the sample."""
    ws, computed = _build([_item(1)])
    q = ws.cell(row=13, column=COL_GST).value
    assert not (isinstance(q, str) and q.startswith("="))
    assert q == computed.lines[0].igst_amount


def test_rate_cells_format_whole_and_fractional_percentages():
    """R/T keep exact values while avoiding rounded or dangling-percent display."""
    ws, _ = _build([_item(1, bcd_rate=0.075, igst_rate=0.18)])

    igst = ws.cell(row=13, column=COL_RATE_OF_DUTY_IGST)
    bcd = ws.cell(row=13, column=COL_RATE_OF_INTEREST_BCD)
    assert igst.value == 0.18
    assert bcd.value == 0.075
    assert igst.number_format == "0%"
    assert bcd.number_format == "0.##%"


def test_usd_rate_integer_renders_without_trailing_zero():
    ws, _ = _build([_item(1)], usd_rate=95.0)
    assert ws.cell(row=13, column=COL_RATE_PER_USD).value == "=L13*95"


def test_totals_row_uses_sum_over_actual_data_rows():
    ws, _ = _build([_item(s) for s in range(1, 4)])  # 3 items -> rows 13..15
    totals_row = 61
    assert ws.cell(row=totals_row, column=COL_AMOUNT).value == "=SUM(L13:L15)"
    assert ws.cell(row=totals_row, column=COL_TOTAL_CUSTOM_DUTY).value == "=SUM(P13:P15)"
    # 'total usd' aux cell mirrors the totals-row L cell.
    assert ws["E71"].value == f"=L{totals_row}"


def test_zero_qty_writes_literal_zero_not_div_formula():
    """Y = O/H would be #DIV/0! at qty==0; the writer emits the literal 0."""
    ws, _ = _build([_item(1, qty=0)])
    assert ws.cell(row=13, column=COL_RATE_PER_UNIT).value == 0


def test_missing_input_leaves_computed_cell_blank_in_formula_mode():
    """A missing unit price must blank the dependent cells (no spurious formula)."""
    ws, _ = _build([_item(1, missing_price=True)])
    # amount (L) depends on the missing price -> blank, not '=H13*K13'.
    assert ws.cell(row=13, column=COL_AMOUNT).value is None
    assert ws.cell(row=13, column=COL_RATE_PER_USD).value is None


def test_literal_mode_emits_numbers_matching_formula_arithmetic():
    items = [_item(1, unit="DOZ", qty=10, price=2)]
    ws_lit, computed = _build(items, use_formulas=False)
    line = computed.lines[0]
    # Literal mode writes the evaluated numbers (no formula strings).
    assert ws_lit.cell(row=13, column=COL_AMOUNT).value == line.amount_usd == 20.0
    assert ws_lit.cell(row=13, column=COL_RATE_PER_USD).value == line.purchase_inr
    assert ws_lit.cell(row=13, column=COL_LAND_COST_WITHOUT_GST).value == line.land_cost_excl_gst
    # Cross-check the formula arithmetic equals the calculator: L=H*K.
    assert line.amount_usd == 10 * 2
