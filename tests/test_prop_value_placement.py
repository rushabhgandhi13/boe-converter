"""Property-based test for full-precision value placement (task 7.4).

Property 10: Every value is written to its mapped cell unchanged at full precision.

**Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 5.3, 5.4, 5.5, 5.6, 5.7, 5.10, 6.11, 6.12, 7.4, 7.5, 8.4, 8.5**

For any document, each extracted value (written verbatim) and each computed value
(written at full floating-point precision) lands in its mapped Header_Block /
Item_Table / Totals_Row cell equal to the source value, with no rounding,
truncation, or reformatting.

Strategy: generate ``ExtractedDocument``s whose header fields and per-line inputs
all parse cleanly to numbers/strings, run them through ``ValueCalculator.compute``
to obtain a ``ComputedDocument``, build the workbook with
``ExcelGenerator.build_workbook``, then assert every mapped cell equals exactly
the corresponding model value:

- Directly-extracted values (parsed verbatim): the cell equals the field's
  ``parsed`` interpretation (``_raw_cell_value`` semantics for clean inputs).
- Computed values: the cell equals the ``ComputedLine`` / ``Totals`` field with
  exact float equality (no rounding).
- Header_Block values at ``D1:G8`` and the Totals_Row package count / sums.
"""

from __future__ import annotations

from hypothesis import given, settings
from hypothesis import strategies as st

from boe_converter.calculator import ValueCalculator
from boe_converter.excel_writer import (
    COL_AMOUNT,
    COL_CTN,
    COL_CUST_AIDC,
    COL_CUSTOM_ASS_VALUE,
    COL_DESCRIPTION,
    COL_GST,
    COL_HSN_CODE,
    COL_LAND_COST_WITH_GST,
    COL_LAND_COST_WITHOUT_GST,
    COL_PCS,
    COL_QTY,
    COL_RATE_OF_DUTY_IGST,
    COL_RATE_OF_INTEREST_BCD,
    COL_RATE_OF_INTEREST_SWS,
    COL_RATE_PER_UNIT,
    COL_RATE_PER_USD,
    COL_SR_NO,
    COL_SURCHARGE,
    COL_TOTAL_CUSTOM_DUTY,
    COL_TOTAL_CUSTOM_DUTY_2,
    COL_UNIT,
    COL_UNIT_PRICE_USD,
    ITEM_TABLE_FIRST_DATA_ROW,
    SWS_RATE,
    TOTALS_ROW,
    ExcelGenerator,
)
from boe_converter.models import (
    ExtractedDocument,
    HeaderBlock,
    LineItem,
    RawValue,
    ReviewFlagSet,
)

# ---------------------------------------------------------------------------
# Generators
# ---------------------------------------------------------------------------

# Bounded, finite floats so the chained per-line products (amount -> purchase
# INR -> land cost ...) stay well clear of float overflow / NaN, letting every
# computed value be compared for exact floating-point equality.
_PRICE = st.floats(min_value=-1e3, max_value=1e3, allow_nan=False, allow_infinity=False)
_QTY = st.floats(min_value=-1e3, max_value=1e3, allow_nan=False, allow_infinity=False)
_MONEY = st.floats(min_value=-1e6, max_value=1e6, allow_nan=False, allow_infinity=False)
_IGST_RATE = st.floats(min_value=0.0, max_value=1.0, allow_nan=False, allow_infinity=False)
_USD_RATE = st.floats(min_value=0.0, max_value=200.0, allow_nan=False, allow_infinity=False)

# Non-empty text that round-trips verbatim through openpyxl (avoid leading '='
# which Excel would treat as a formula; avoid pure control characters).
_TEXT = st.text(
    alphabet=st.characters(min_codepoint=32, max_codepoint=126, blacklist_characters="="),
    min_size=1,
    max_size=20,
)


def _num(value: float) -> RawValue:
    """A located, cleanly-parseable numeric field, as the parser emits."""

    return RawValue(raw_text=repr(value), parsed=value)


def _str(value: str) -> RawValue:
    """A located, cleanly-parseable text field, as the parser emits."""

    return RawValue(raw_text=value, parsed=value)


@st.composite
def _line_items(draw):
    """A list of LineItems with unique, ascending serials and clean inputs."""

    count = draw(st.integers(min_value=0, max_value=12))
    items: list[LineItem] = []
    for serial in range(1, count + 1):
        items.append(
            LineItem(
                item_serial=serial,
                cth_hsn=_str(draw(_TEXT)),
                description=_str(draw(_TEXT)),
                unit_price_usd=_num(draw(_PRICE)),
                quantity=_num(draw(_QTY)),
                unit=_str(draw(st.sampled_from(("DOZ", "PCS", "KGS", "SET", " doz ")))),
                assessable_value=_num(draw(_MONEY)),
                bcd_rate=_num(draw(_MONEY)),
                bcd_amount=_num(draw(_MONEY)),
                igst_rate=_num(draw(_IGST_RATE)),
                total_duty=_num(draw(_MONEY)),
            )
        )
    # Shuffle so the writer's ascending-serial sort is genuinely exercised.
    draw(st.randoms(use_true_random=False)).shuffle(items)
    return items


@st.composite
def _documents(draw):
    """An ExtractedDocument with a fully-populated, cleanly-parseable header."""

    usd_rate = draw(_USD_RATE)
    header = HeaderBlock(
        company_name=draw(_TEXT),
        party_name=_str(draw(_TEXT)),
        usd_rate=usd_rate,
        details=_str(draw(_TEXT)),
        invoice_no=_str(draw(_TEXT)),
        invoice_date=_str(draw(_TEXT)),
        be_no=_str(draw(_TEXT)),
        be_date=_str(draw(_TEXT)),
        bl_no=_str(draw(_TEXT)),
        bl_date=_str(draw(_TEXT)),
        invoice_amount=_num(draw(_MONEY)),
        invoice_currency=_str("USD"),
        package_count=_num(draw(st.floats(
            min_value=0, max_value=1e6, allow_nan=False, allow_infinity=False
        ))),
        container_details=_str(draw(_TEXT)),
    )
    doc = ExtractedDocument(
        header=header,
        line_items=draw(_line_items()),
        declared_item_count=None,
        flags=[],
    )
    return doc, usd_rate


def _sum_optional(values) -> float:
    """Sum float|None values skipping None, in iteration order (mirrors writer)."""

    total = 0.0
    for value in values:
        if value is not None:
            total += value
    return total


# ---------------------------------------------------------------------------
# Property 10
# ---------------------------------------------------------------------------


@given(payload=_documents())
@settings(deadline=None, max_examples=25)
def test_values_placed_in_mapped_cells_at_full_precision(payload):
    """Every extracted/computed value lands in its mapped cell unchanged."""

    doc, usd_rate = payload
    computed = ValueCalculator().compute(doc, usd_rate)
    # Literal mode: assert the evaluated computed values land in their mapped
    # cells. The live-formula output is covered by tests/test_excel_formulas.py.
    ws = ExcelGenerator(use_formulas=False).build_workbook(computed, ReviewFlagSet()).active

    # --- Header_Block (D1:G8) -------------------------------------------
    header = computed.header
    assert ws["E1"].value == header.company_name          # Req 4.7
    assert ws["E2"].value == header.party_name.parsed      # Req 4.1
    assert ws["G2"].value == header.usd_rate               # Req 4.5
    assert ws["E3"].value == header.details.parsed         # Req 4.6
    assert ws["E4"].value == header.invoice_no.parsed      # Req 4.2
    assert ws["G4"].value == header.invoice_date.parsed    # Req 4.2
    assert ws["E5"].value == header.be_no.parsed           # Req 4.3
    assert ws["G5"].value == header.be_date.parsed         # Req 4.3
    assert ws["E6"].value == header.bl_no.parsed           # Req 4.4
    assert ws["G6"].value == header.bl_date.parsed         # Req 4.4

    # --- Item_Table data rows (ascending serial from row 13) ------------
    ordered = sorted(computed.lines, key=lambda ln: ln.source.item_serial)
    for offset, line in enumerate(ordered):
        row = ITEM_TABLE_FIRST_DATA_ROW + offset
        item = line.source

        def cell(col):
            return ws.cell(row=row, column=col).value

        # Sr. no. - dense sequential 1..N (Req 5.2).
        assert cell(COL_SR_NO) == offset + 1

        # Directly-extracted values, verbatim parsed value (Req 5.3-5.10, 8.4).
        assert cell(COL_DESCRIPTION) == item.description.parsed       # Req 5.3
        assert cell(COL_HSN_CODE) == item.cth_hsn.parsed              # Req 5.4
        assert cell(COL_QTY) == item.quantity.parsed                  # Req 5.6
        assert cell(COL_UNIT) == item.unit.parsed                     # Req 5.7
        assert cell(COL_UNIT_PRICE_USD) == item.unit_price_usd.parsed  # Req 5.5
        assert cell(COL_CUSTOM_ASS_VALUE) == item.assessable_value.parsed  # Req 5.10
        assert cell(COL_RATE_OF_DUTY_IGST) == item.igst_rate.parsed   # Req 6.11
        assert cell(COL_RATE_OF_INTEREST_BCD) == item.bcd_rate.parsed  # Req 6.11
        assert cell(COL_CUST_AIDC) == item.bcd_amount.parsed          # Req 6.11
        assert cell(COL_RATE_OF_INTEREST_SWS) == SWS_RATE             # constant 0.10

        # Computed values, full precision / exact float equality (Req 6.12, 8.5).
        assert cell(COL_PCS) == line.pcs
        assert cell(COL_AMOUNT) == line.amount_usd
        assert cell(COL_RATE_PER_USD) == line.purchase_inr
        assert cell(COL_LAND_COST_WITHOUT_GST) == line.land_cost_excl_gst
        assert cell(COL_TOTAL_CUSTOM_DUTY) == line.total_customs_duty
        assert cell(COL_GST) == line.igst_amount
        assert cell(COL_TOTAL_CUSTOM_DUTY_2) == line.combined_duty
        assert cell(COL_SURCHARGE) == line.sws_amount
        assert cell(COL_LAND_COST_WITH_GST) == line.land_cost_incl_gst
        assert cell(COL_RATE_PER_UNIT) == line.purchase_rate_per_unit

    # --- Totals_Row (row 61): full-precision sums + package count -------
    totals = computed.totals

    def total_cell(col):
        return ws.cell(row=TOTALS_ROW, column=col).value

    # G: BOE total package/CTN count, carried through verbatim (Req 7.5).
    assert total_cell(COL_CTN) == totals.package_count.parsed
    # Columns backed by pre-computed Totals fields (Req 7.4, 8.5).
    assert total_cell(COL_AMOUNT) == totals.total_amount_usd
    assert total_cell(COL_CUSTOM_ASS_VALUE) == totals.total_assessable_value
    assert total_cell(COL_LAND_COST_WITHOUT_GST) == totals.total_land_cost_excl_gst
    assert total_cell(COL_TOTAL_CUSTOM_DUTY) == totals.total_customs_duty
    assert total_cell(COL_GST) == totals.total_igst
    assert total_cell(COL_LAND_COST_WITH_GST) == totals.total_land_cost_incl_gst
    # Columns summed in the writer from per-line values (full precision).
    assert total_cell(COL_RATE_PER_USD) == _sum_optional(
        ln.purchase_inr for ln in computed.lines
    )
    assert total_cell(COL_TOTAL_CUSTOM_DUTY_2) == _sum_optional(
        ln.combined_duty for ln in computed.lines
    )
    assert total_cell(COL_CUST_AIDC) == _sum_optional(
        float(ln.source.bcd_amount.parsed) for ln in computed.lines
    )
    assert total_cell(COL_SURCHARGE) == _sum_optional(
        ln.sws_amount for ln in computed.lines
    )
