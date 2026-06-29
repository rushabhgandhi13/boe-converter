"""Property-based test for no-source and sample-empty cells (Excel_Generator).

Property 11: No-source and sample-empty cells are always empty.

**Validates: Requirements 4.8, 4.9, 8.6**

For any document, the workbook produced by ``ExcelGenerator.build_workbook``
must leave the following cells truly empty (``value is None`` -> no characters,
spaces, or placeholder text):

- **Header no-source value cells** (Req 4.8): the ``USD Amt`` value (G3), the
  ``Eway bill no`` / ``Eway bill date`` values (E7 / G7), and the
  ``RETTENCE DATE`` / ``RETTENCE RATE`` values (E8 / G8). Only the labels for
  these fields are written; their value cells never carry a value.
- **Sample-empty Item_Table columns** for every data row (Req 8.6): ``PARTY
  NAME`` (B), ``BILLING AMOUNT`` (C), ``AS PER TALLY NAME`` (D) and the
  per-row ``CTN`` (G) column, which is Manual/External in Milestone 1.
- **Flagged/missing-source cells** (Req 4.9): any directly-mapped Line_Item or
  Header_Block cell whose source ``RawValue`` was flagged missing
  (``RawValue.missing()``) must be left blank rather than filled with an
  inferred or placeholder value.

Strategy: generate varied ``ComputedDocument``s (built through the real
``ValueCalculator`` from generated ``LineItem``s and a ``HeaderBlock`` whose
fields are independently present or missing), build the workbook, and assert the
above cells are ``None``.
"""

from __future__ import annotations

from hypothesis import given, settings
from hypothesis import strategies as st

from boe_converter.calculator import ValueCalculator
from boe_converter.excel_writer import (
    COL_AS_PER_TALLY_NAME,
    COL_BILLING_AMOUNT,
    COL_CTN,
    COL_CUSTOM_ASS_VALUE,
    COL_CUST_AIDC,
    COL_DESCRIPTION,
    COL_HSN_CODE,
    COL_PARTY_NAME,
    COL_QTY,
    COL_RATE_OF_DUTY_IGST,
    COL_RATE_OF_INTEREST_BCD,
    COL_UNIT,
    COL_UNIT_PRICE_USD,
    ITEM_TABLE_FIRST_DATA_ROW,
    ExcelGenerator,
)
from boe_converter.models import (
    ExtractedDocument,
    HeaderBlock,
    LineItem,
    RawValue,
    ReviewFlagSet,
)

# --- Generators ------------------------------------------------------------

_FINITE = st.floats(
    min_value=-1e9,
    max_value=1e9,
    allow_nan=False,
    allow_infinity=False,
)


def _numeric_raw(value: float) -> RawValue:
    """A located, parseable numeric field as it would arrive from the parser."""

    return RawValue(raw_text=str(value), parsed=value)


def _text_raw(value: str) -> RawValue:
    """A located, parseable text field."""

    return RawValue(raw_text=value, parsed=value)


# Printable, control-char-free text. openpyxl rejects ASCII control characters
# in cell values; the parser would never emit them, so we keep generated text
# within a safe printable alphabet to exercise the writer realistically.
_SAFE_TEXT = st.text(
    alphabet=st.characters(min_codepoint=32, max_codepoint=126),
    min_size=1,
    max_size=12,
)


# A numeric field that is either present (numeric) or flagged missing. We focus
# on present-vs-missing (rather than unparseable) because Property 11 concerns
# cells that must be *empty*: a missing source blanks the cell, whereas an
# unparseable source intentionally writes its raw text (Req 9.2) and so is out
# of scope for the "always empty" property.
_MAYBE_MISSING_NUMERIC = st.one_of(
    _FINITE.map(_numeric_raw),
    st.just(RawValue.missing()),
)

_MAYBE_MISSING_TEXT = st.one_of(
    _SAFE_TEXT.map(_text_raw),
    st.just(RawValue.missing()),
)

_UNIT = st.one_of(
    st.sampled_from(("DOZ", "PCS", "KGS", "SET", " doz ")).map(_text_raw),
    st.just(RawValue.missing()),
)

_LINE_ITEMS = st.builds(
    LineItem,
    item_serial=st.integers(min_value=1, max_value=10_000),
    cth_hsn=_MAYBE_MISSING_TEXT,
    description=_MAYBE_MISSING_TEXT,
    unit_price_usd=_MAYBE_MISSING_NUMERIC,
    quantity=_MAYBE_MISSING_NUMERIC,
    unit=_UNIT,
    assessable_value=_MAYBE_MISSING_NUMERIC,
    bcd_rate=_MAYBE_MISSING_NUMERIC,
    bcd_amount=_MAYBE_MISSING_NUMERIC,
    igst_rate=_MAYBE_MISSING_NUMERIC,
    total_duty=_MAYBE_MISSING_NUMERIC,
)


def _maybe_missing_pkg() -> st.SearchStrategy[RawValue]:
    return st.one_of(
        st.integers(min_value=0, max_value=100_000).map(
            lambda n: RawValue(raw_text=str(n), parsed=n)
        ),
        st.just(RawValue.missing()),
    )


_HEADER = st.builds(
    HeaderBlock,
    company_name=st.sampled_from(("Gemini Unicom LLP", "")),
    party_name=_MAYBE_MISSING_TEXT,
    usd_rate=_FINITE,
    details=_MAYBE_MISSING_TEXT,
    invoice_no=_MAYBE_MISSING_TEXT,
    invoice_date=_MAYBE_MISSING_TEXT,
    be_no=_MAYBE_MISSING_TEXT,
    be_date=_MAYBE_MISSING_TEXT,
    bl_no=_MAYBE_MISSING_TEXT,
    bl_date=_MAYBE_MISSING_TEXT,
    invoice_amount=_MAYBE_MISSING_NUMERIC,
    invoice_currency=_MAYBE_MISSING_TEXT,
    package_count=_maybe_missing_pkg(),
    container_details=_MAYBE_MISSING_TEXT,
)


# Header_Block value cells that have no BOE/configuration source (Req 4.8): the
# label is written but the value cell must always stay empty.
_HEADER_NO_SOURCE_VALUE_CELLS = ("G3", "E7", "G7", "E8", "G8")

# Item_Table columns left empty for every data row (Req 8.6 / Milestone-1
# Manual/External): PARTY NAME (B), BILLING AMOUNT (C), AS PER TALLY NAME (D),
# and the per-row CTN (G).
_ROW_EMPTY_COLUMNS = (
    COL_PARTY_NAME,
    COL_BILLING_AMOUNT,
    COL_AS_PER_TALLY_NAME,
    COL_CTN,
)

# Direct Line_Item fields -> their mapped Item_Table column. A field whose
# source RawValue is flagged missing must leave its mapped cell empty (Req 4.9).
_DIRECT_FIELD_COLUMNS = {
    "description": COL_DESCRIPTION,
    "cth_hsn": COL_HSN_CODE,
    "quantity": COL_QTY,
    "unit": COL_UNIT,
    "unit_price_usd": COL_UNIT_PRICE_USD,
    "assessable_value": COL_CUSTOM_ASS_VALUE,
    "igst_rate": COL_RATE_OF_DUTY_IGST,
    "bcd_rate": COL_RATE_OF_INTEREST_BCD,
    "bcd_amount": COL_CUST_AIDC,
}


def _build_workbook(header: HeaderBlock, items: list[LineItem]):
    """Run items + header through the calculator, then build the workbook."""

    calc = ValueCalculator()
    extracted = ExtractedDocument(header=header, line_items=items)
    computed = calc.compute(extracted, header.usd_rate)
    flags = computed.flags
    wb = ExcelGenerator().build_workbook(computed, ReviewFlagSet(flags))
    return wb["Sheet1"], computed


@given(header=_HEADER, items=st.lists(_LINE_ITEMS, max_size=15))
@settings(deadline=None, max_examples=25)
def test_header_no_source_value_cells_are_empty(header, items):
    """The no-source Header_Block value cells (Req 4.8) carry no value."""

    ws, _ = _build_workbook(header, items)

    for coordinate in _HEADER_NO_SOURCE_VALUE_CELLS:
        assert ws[coordinate].value is None, (
            f"no-source header cell {coordinate} should be empty"
        )


@given(header=_HEADER, items=st.lists(_LINE_ITEMS, min_size=1, max_size=15))
@settings(deadline=None, max_examples=25)
def test_sample_empty_columns_are_blank_for_every_data_row(header, items):
    """Columns B/C/D/G are empty for every line-item data row (Req 8.6)."""

    ws, computed = _build_workbook(header, items)

    n = len(computed.lines)
    for offset in range(n):
        row = ITEM_TABLE_FIRST_DATA_ROW + offset
        for col in _ROW_EMPTY_COLUMNS:
            assert ws.cell(row=row, column=col).value is None, (
                f"sample-empty column {col} at row {row} should be blank"
            )


@given(header=_HEADER, items=st.lists(_LINE_ITEMS, min_size=1, max_size=15))
@settings(deadline=None, max_examples=25)
def test_missing_source_fields_leave_their_cells_blank(header, items):
    """A field flagged missing leaves its mapped Item_Table cell empty (Req 4.9)."""

    ws, computed = _build_workbook(header, items)

    ordered = sorted(computed.lines, key=lambda ln: ln.source.item_serial)
    for offset, line in enumerate(ordered):
        row = ITEM_TABLE_FIRST_DATA_ROW + offset
        item = line.source
        for field_name, col in _DIRECT_FIELD_COLUMNS.items():
            rv = getattr(item, field_name)
            if rv.is_missing:
                assert ws.cell(row=row, column=col).value is None, (
                    f"missing field {field_name} at row {row} (col {col}) "
                    "should be blank"
                )
