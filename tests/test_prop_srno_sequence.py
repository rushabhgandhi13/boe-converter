"""Property-based test for the Sr. no. sequence and dense, ascending rows.

Property 9: Sr. no. forms the consecutive sequence 1..N and rows are dense and
ascending.

**Validates: Requirements 5.1, 5.2**

Strategy: generate ``N`` line items with arbitrary, distinct item serials in a
shuffled order, wrap each in a ``ComputedLine`` (encoding the serial into the
Description cell so the written row order can be recovered), assemble a
``ComputedDocument`` and build the workbook. Reading column ``A`` from row 13
downward must yield exactly the sequence ``1, 2, ..., N`` with no gaps, repeats,
or blank rows, there must be exactly ``N`` populated data rows, and the items
written into those rows must appear in ascending item-serial order beginning at
row 13.
"""

from __future__ import annotations

from hypothesis import given, settings
from hypothesis import strategies as st

from boe_converter.excel_writer import (
    COL_DESCRIPTION,
    COL_SR_NO,
    ITEM_TABLE_FIRST_DATA_ROW,
    ExcelGenerator,
)
from boe_converter.models import (
    ComputedDocument,
    ComputedLine,
    HeaderBlock,
    LineItem,
    RawValue,
    ReviewFlagSet,
    Totals,
)


def _raw(value) -> RawValue:
    """A located, parseable field as it would arrive from the parser."""

    return RawValue(raw_text=str(value), parsed=value)


def _header() -> HeaderBlock:
    """A minimal, fully-present header (its content is irrelevant to Property 9)."""

    return HeaderBlock(
        company_name="Gemini Unicom LLP",
        party_name=_raw("ACME"),
        usd_rate=90.0,
        details=_raw("CO-32"),
        invoice_no=_raw("INV-1"),
        invoice_date=_raw("01/01/2026"),
        be_no=_raw("BE-1"),
        be_date=_raw("01/01/2026"),
        bl_no=_raw("BL-1"),
        bl_date=_raw("01/01/2026"),
        invoice_amount=_raw(100.0),
        invoice_currency=_raw("USD"),
        package_count=_raw(1357),
        container_details=_raw("CONT-1"),
    )


def _line(serial: int) -> ComputedLine:
    """A ComputedLine whose source encodes its serial in the Description cell.

    Description is written into column E by the writer, so reading column E back
    lets the test confirm the rows were emitted in ascending item-serial order.
    """

    item = LineItem(
        item_serial=serial,
        cth_hsn=_raw("1234"),
        description=_raw(serial),  # parsed int -> written to column E verbatim
        unit_price_usd=_raw(1.0),
        quantity=_raw(1.0),
        unit=_raw("PCS"),
        assessable_value=_raw(1.0),
        bcd_rate=_raw(0.0),
        bcd_amount=_raw(0.0),
        igst_rate=_raw(0.0),
        total_duty=_raw(0.0),
    )
    return ComputedLine(source=item)


# Distinct serials (the join key is unique per item), shuffled into arbitrary
# order so the writer must sort them; kept modest in count to stay within the
# sample's data region and well clear of the fixed totals/auxiliary rows.
_SERIAL_LISTS = st.lists(
    st.integers(min_value=1, max_value=10_000),
    min_size=1,
    max_size=40,
    unique=True,
)


@given(serials=_SERIAL_LISTS)
@settings(deadline=None)
def test_sr_no_is_dense_consecutive_and_rows_ascending(serials):
    """Sr. no. == 1..N, exactly N dense rows, items in ascending serial order."""

    n = len(serials)
    lines = [_line(s) for s in serials]
    doc = ComputedDocument(header=_header(), lines=lines, totals=Totals())

    wb = ExcelGenerator().build_workbook(doc, ReviewFlagSet())
    ws = wb["Sheet1"]

    first = ITEM_TABLE_FIRST_DATA_ROW

    # Column A over the N data rows is exactly the sequence 1, 2, ..., N.
    sr_values = [
        ws.cell(row=first + offset, column=COL_SR_NO).value for offset in range(n)
    ]
    assert sr_values == list(range(1, n + 1))

    # No gaps/blanks within the data rows: every Sr. no. cell is populated.
    assert all(value is not None for value in sr_values)

    # Exactly N data rows: the row immediately after the block has no Sr. no.
    assert ws.cell(row=first + n, column=COL_SR_NO).value is None

    # The items occupy the rows in ascending item-serial order from row 13.
    written_serials = [
        ws.cell(row=first + offset, column=COL_DESCRIPTION).value
        for offset in range(n)
    ]
    assert written_serials == sorted(serials)
