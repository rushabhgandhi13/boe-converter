"""Golden-file unit tests pinning the generator's labels to the sample workbook.

Task 7.7. The Excel_Generator reproduces the CTN layout's fixed text exactly:
the Item_Table header row (row 12, columns A..Y) and the auxiliary template
section labels (``DETAILS AS PER CHALLANS`` / ``DETAILS AS PER TALLY`` /
``CLEARANCE AND FORWARDING INVOICE`` titles, their column-header rows, the C&F
detail block row labels, and the ``total usd`` summary label) must each match
the real golden workbook ``1357 ctn llp.xlsx`` character-for-character,
including deliberate quirks (trailing/double spaces, typos like
``'Rate Per USDin purcahse '`` and ``'EWAY BILL N0'``).

These tests open the actual golden workbook with openpyxl and compare, cell by
cell, the captured label constants in :mod:`boe_converter.excel_writer`
(``ITEM_TABLE_HEADERS`` keyed by column index on row 12, and ``AUX_LABELS``
keyed by cell coordinate) against the workbook's own cell values. This guards
against any drift between the captured constants and the file they were taken
from (Req 5.12, 8.3).

The golden workbook is a project asset that may be absent in some checkouts, so
the whole module is skipped when it is not present.

Validates: Requirements 5.12, 8.3
"""

from __future__ import annotations

from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")

from boe_converter.excel_writer import (  # noqa: E402
    AUX_LABELS,
    ITEM_TABLE_HEADER_ROW,
    ITEM_TABLE_HEADERS,
)

# The golden sample workbook lives at the project root (one level above tests/).
GOLDEN_WORKBOOK = Path(__file__).resolve().parent.parent / "1357 ctn llp.xlsx"

# Module-level skip when the golden asset is not available.
pytestmark = pytest.mark.skipif(
    not GOLDEN_WORKBOOK.exists(),
    reason=f"golden sample workbook not found at {GOLDEN_WORKBOOK}",
)


@pytest.fixture(scope="module")
def golden_sheet():
    """Open ``Sheet1`` of the golden workbook once (formulas preserved)."""
    wb = openpyxl.load_workbook(str(GOLDEN_WORKBOOK), data_only=False)
    try:
        yield wb["Sheet1"]
    finally:
        wb.close()


@pytest.mark.parametrize(
    "column, label",
    sorted(ITEM_TABLE_HEADERS.items()),
    ids=[f"col{col}" for col in sorted(ITEM_TABLE_HEADERS)],
)
def test_item_table_header_matches_golden(golden_sheet, column, label):
    """Req 5.12 - each Item_Table header label (row 12, col A..Y) equals the
    golden workbook's cell value character-for-character."""
    actual = golden_sheet.cell(row=ITEM_TABLE_HEADER_ROW, column=column).value
    assert actual == label, (
        f"row {ITEM_TABLE_HEADER_ROW} col {column}: captured {label!r} != "
        f"sample {actual!r}"
    )


@pytest.mark.parametrize(
    "coordinate, label",
    sorted(AUX_LABELS.items()),
    ids=sorted(AUX_LABELS),
)
def test_aux_label_matches_golden(golden_sheet, coordinate, label):
    """Req 8.3 - each auxiliary template label equals the golden workbook's cell
    value at that coordinate character-for-character (quirks preserved)."""
    actual = golden_sheet[coordinate].value
    assert actual == label, (
        f"{coordinate}: captured {label!r} != sample {actual!r}"
    )


def test_item_table_headers_cover_all_columns():
    """The captured header map covers exactly columns A..Y (1..25) with no gaps,
    matching the golden Item_Table header span (Req 5.12)."""
    assert sorted(ITEM_TABLE_HEADERS) == list(range(1, 26))
