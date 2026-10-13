"""Excel_Generator: writes the ComputedDocument into the exact CTN layout.

Milestone 1. Builds ``Sheet1`` in the exact ``1357 ctn llp.xlsx`` layout:

- Header_Block in ``D1:G8`` (labels in column D/F, values in column E/G).
- Item_Table header at row 12 (``A12:Y12``), with each label reproduced
  character-for-character from the sample (including its quirks, e.g.
  ``'Rate Per USDin purcahse '`` and ``'Ratepurchase per unitPcs/KGS/SET'``).
- Line items written from row 13, one dense row per item in ascending serial
  order with no blank rows, ``Sr. no.`` running 1..N.

Computed values are written as full-precision literals (Req 8.5); directly
extracted values are written verbatim (Req 8.4). Review-flagged / missing cells
and no-source columns are left blank (Req 5.11, 6.13, 4.8, 4.9, 8.6); a field
that was located but could not be parsed is written as its raw text (Req 9.2).

The Totals_Row (row 61) and the auxiliary template sections are written by
``_write_totals_row`` / ``_write_aux_templates`` (task 7.2). The totals row
reproduces the sample's summed columns (``G, L, M, N, O, P, Q, S, U, W, X``)
plus the BOE package count in column G; the auxiliary sections reproduce the
``DETAILS AS PER CHALLANS`` / ``DETAILS AS PER TALLY`` / ``CLEARANCE AND
FORWARDING INVOICE`` and C&F block labels verbatim with empty data cells.

Cell map grounded directly in the sample workbook (design.md "Target Excel cell
map"):

    A Sr. no.      (computed 1..N)        N CUSTOM ASS VALUE            (direct)
    B PARTY NAME   (blank)                O LAND COST ... WITHOUT GST   (computed)
    C BILLING AMOUNT (blank)              P TOTAL Custom Duty           (computed)
    D AS PER TALLY NAME (blank)           Q GST                         (computed)
    E Description  (direct)               R RATE OF DUTY IGST           (direct rate)
    F HSN CODE     (direct)               S total custom duty           (computed)
    G CTN          (blank)                T RATE OF INTEREST            (direct BCD rate)
    H QTY          (direct)               U CUST AIDC                   (direct BCD amount)
    I Unit         (direct)               V RATE OF INTEREST            (constant 0.10 SWS)
    J pcs          (computed)             W SURCHARGE                   (computed SWS)
    K Unit Price in USD (direct)          X LAND COST ... WITH GST      (computed)
    L Amount       (computed)             Y Ratepurchase per unit...    (computed)
    M Rate Per USDin purcahse  (computed)
"""

from __future__ import annotations

from copy import copy
from functools import lru_cache
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from boe_converter.models import (
    ComputedDocument,
    ComputedLine,
    GrandTotalVerification,
    HeaderBlock,
    RawValue,
    ReviewFlagSet,
    Totals,
)

# The bundled golden workbook used as the *style template*: it carries the
# sample's exact column widths, row heights, merged ranges, fonts, fills,
# borders, number formats and the static decorative band / auxiliary
# scaffolding. The generator loads it as the base workbook, clears the
# document-specific cells, and writes this document's values into the
# already-styled cells, so the output is visually identical to
# ``1357 ctn llp.xlsx`` (Req 8.x). Only data is replaced; styling is inherited.
_TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "ctn_template.xlsx"

# Template layout anchors (the sample's fixed structure).
_TEMPLATE_FIRST_DATA_ROW = 13
_TEMPLATE_LAST_DATA_ROW = 60          # styled data region rows 13..60
_TEMPLATE_TOTALS_ROW = 61
# Blank item rows left between the last line item and the Totals_Row, so the
# total always sits two rows below the data (e.g. 18 items in rows 13..30 put
# the totals at row 33) instead of at a fixed row with a gap of empty rows.
BLANK_ROWS_AFTER_ITEMS = 2
# A representative interior data row whose per-column styling (borders, number
# formats, fonts) is cloned onto continuation rows when a document has more
# line items than the template's styled data region (overflow handling).
_TEMPLATE_STYLE_ROW = 14


@lru_cache(maxsize=1)
def _template_bytes() -> bytes:
    """Read and cache the template workbook bytes (avoids repeated disk I/O)."""
    return _TEMPLATE_PATH.read_bytes()


def _load_template() -> Workbook:
    """Return a fresh, writable copy of the style template workbook."""
    return load_workbook(BytesIO(_template_bytes()))

# --- Item_Table column indices (1-based, A=1 .. Y=25) ----------------------
COL_SR_NO = 1                       # A  Sr. no.            (computed sequential)
COL_PARTY_NAME = 2                  # B  PARTY NAME         (blank, no source)
COL_BILLING_AMOUNT = 3              # C  BILLING AMOUNT     (blank, no source)
COL_AS_PER_TALLY_NAME = 4           # D  AS PER TALLY NAME  (blank, Manual/External)
COL_DESCRIPTION = 5                 # E  Description        (direct)
COL_HSN_CODE = 6                    # F  HSN CODE           (direct)
COL_CTN = 7                         # G  CTN                (blank, Manual/External)
COL_QTY = 8                         # H  QTY                (direct)
COL_UNIT = 9                        # I  Unit               (direct)
COL_PCS = 10                        # J  pcs                (computed)
COL_UNIT_PRICE_USD = 11            # K  Unit Price in USD   (direct)
COL_AMOUNT = 12                     # L  Amount             (computed)
COL_RATE_PER_USD = 13              # M  Rate Per USDin purcahse  (computed)
COL_CUSTOM_ASS_VALUE = 14         # N  CUSTOM ASS VALUE    (direct)
COL_LAND_COST_WITHOUT_GST = 15    # O  LAND COST ... WITHOUT GST (computed)
COL_TOTAL_CUSTOM_DUTY = 16        # P  TOTAL Custom Duty   (computed)
COL_GST = 17                       # Q  GST                (computed)
COL_RATE_OF_DUTY_IGST = 18        # R  RATE OF DUTY IGST   (direct igst rate)
COL_TOTAL_CUSTOM_DUTY_2 = 19      # S  total custom duty   (computed combined duty)
COL_RATE_OF_INTEREST_BCD = 20     # T  RATE OF INTEREST    (direct BCD rate)
COL_CUST_AIDC = 21                 # U  CUST AIDC          (direct BCD amount)
COL_RATE_OF_INTEREST_SWS = 22     # V  RATE OF INTEREST    (constant 0.10 SWS rate)
COL_SURCHARGE = 23                 # W  SURCHARGE          (computed SWS amount)
COL_LAND_COST_WITH_GST = 24       # X  LAND COST ... WITH GST (computed)
COL_RATE_PER_UNIT = 25            # Y  Ratepurchase per unitPcs/KGS/SET (computed)

ITEM_TABLE_HEADER_ROW = 12
ITEM_TABLE_FIRST_DATA_ROW = 13

# The Totals_Row's template position (row 61 in the sample). The actual totals
# row floats: two blank rows below the last line item
# (``_totals_row_for(n_items)``), so a short document has no bank of empty rows
# and a long one still never collides with the items. ``shift`` throughout the
# writer is the signed displacement of the actual totals row from this anchor
# (negative when rows were removed, positive on overflow).
TOTALS_ROW = 61


def _totals_row_for(n_items: int) -> int:
    """Return the Totals_Row for ``n_items`` line items.

    Items occupy rows 13..(12+N); two blank rows follow; the totals sit below
    them (row 15+N). An empty document still leaves its two blanks (row 15).
    """
    return _TEMPLATE_FIRST_DATA_ROW + max(n_items, 0) + BLANK_ROWS_AFTER_ITEMS

# The constant SWS rate written into column V for every line item (sample V=0.1).
SWS_RATE = 0.10

# Maps a grand-total verification column letter (Req 11) to the Totals_Row
# column index whose cell is red-filled on a MISMATCH: P -> TOTAL Custom Duty,
# Q -> GST, S -> total custom duty.
GRAND_TOTAL_FILL_COLUMNS: dict[str, int] = {
    "P": COL_TOTAL_CUSTOM_DUTY,
    "Q": COL_GST,
    "S": COL_TOTAL_CUSTOM_DUTY_2,
}

# Solid red fill applied to a Totals_Row cell whose grand total mismatches the
# BOE declared value (Req 11.4). MATCH/UNVERIFIED columns stay unstyled.
GRAND_TOTAL_MISMATCH_FILL = PatternFill(
    fill_type="solid", fgColor="FFFF0000", start_color="FFFF0000"
)

# Solid red fill applied to every cell of an item row whose BOE quantity, unit
# price or amount differs from the supplier invoice (see
# ``LineItem.invoice_mismatch``). Same red as a grand-total mismatch: either
# way the operator must look at the row before booking it. Rows that agree -
# or were never compared (no invoice supplied) - keep the template styling.
INVOICE_MISMATCH_FILL = PatternFill(
    fill_type="solid", fgColor="FFFF0000", start_color="FFFF0000"
)

# ---------------------------------------------------------------------------
# Auxiliary template labels, captured character-for-character from the sample
# workbook (``1357 ctn llp.xlsx``). Keyed by cell coordinate -> verbatim label.
# Trailing spaces, double spaces and quirks (e.g. 'EWAY BILL N0' with a zero,
# 'BL  NO. ' with two spaces) are deliberate and must be preserved (Req 8.3).
# Only the labels are written; the data cells within these sections are left
# empty in Milestone 1 (data entry / Tally linkage deferred to Milestone B).
# ---------------------------------------------------------------------------
AUX_LABELS: dict[str, str] = {
    # 'total usd' summary label (its value cell E71 is written separately as the
    # computed total USD, mirroring the sample's '=L61').
    "B71": "total usd",
    # DETAILS AS PER CHALLANS section title + column header row (row 72).
    "G71": "DETAILS AS PER CHALLANS",
    "G72": "MARKA",
    "H72": "PARTY NAME",
    "I72": "CTN",
    "J72": "DATE ",
    "K72": "CHALLAN NO.",
    "L72": "TAXABLE",
    "M72": "GST VALUE",
    "N72": "TOTAL CHALLAN AMT.",
    "O72": "GST RATE",
    # DETAILS AS PER TALLY section title + column header row (row 72).
    "Q71": "DETAILS AS PER TALLY",
    "Q72": "SR NO",
    "R72": "PARTY NAME",
    "S72": "DATE ",
    "T72": "BILL NO",
    "U72": "TAXABLE VALUE",
    "V72": "GST VALUE",
    "W72": "ROUND OFF",
    "X72": "TOTAL VALUE",
    # C&F detail block - row labels in column B (B72-B86); data cells (E*) empty.
    "B72": "customer and other expenses",
    "B73": "IGST",
    "B74": "C&F AGENCY",
    "B75": "B/E NO",
    "B76": "B/E DATE",
    "B77": "BL  NO. ",
    "B78": "BL DATE",
    "B79": "NET WEIGHT",
    "B80": "GROSS WEIGHT",
    "B81": "agst bill no ",
    "B82": "INVOICE RATE  IN USD",
    "B83": "EWAY BILL N0",
    "B84": "CHA CO NAME",
    "B85": "remdince rate",
    "B86": "remdince date",
    # CLEARANCE AND FORWARDING INVOICE section title + column header row (row 95).
    "G94": "CLEARANCE AND FORWARDING INVOICE",
    "G95": "SR NO",
    "H95": "PARTY NAME",
    "I95": "DATE ",
    "J95": "INVOICE NO",
    "K95": "NET AMOUNT",
    "L95": "GST VALUE",
    "M95": "TDS",
    "N95": "ROUND OFF",
    "O95": "TOTALAMOUNT",
}

# The 'total usd' value cell (B71 label -> E71 value), mirroring sample '=L61'.
AUX_TOTAL_USD_CELL = "E71"

# Item_Table header labels, reproduced character-for-character from the sample
# (Req 5.12). Keyed by 1-based column index. Spacing/typo quirks are deliberate.
ITEM_TABLE_HEADERS: dict[int, str] = {
    COL_SR_NO: "Sr. no.",
    COL_PARTY_NAME: "PARTY NAME",
    COL_BILLING_AMOUNT: "BILLING AMOUNT",
    COL_AS_PER_TALLY_NAME: "AS PER TALLY NAME",
    COL_DESCRIPTION: "Description",
    COL_HSN_CODE: "HSN CODE",
    COL_CTN: "CTN",
    COL_QTY: "QTY",
    COL_UNIT: "Unit",
    COL_PCS: "pcs",
    COL_UNIT_PRICE_USD: "Unit Price in USD",
    COL_AMOUNT: "Amount",
    COL_RATE_PER_USD: "Rate Per USDin purcahse ",
    COL_CUSTOM_ASS_VALUE: "CUSTOM ASS VALUE",
    COL_LAND_COST_WITHOUT_GST: "LAND COST OF PURCHASE WITHOUT GST",
    COL_TOTAL_CUSTOM_DUTY: "TOTAL Custom Duty",
    COL_GST: "GST",
    COL_RATE_OF_DUTY_IGST: "RATE OF DUTY IGST",
    COL_TOTAL_CUSTOM_DUTY_2: "total custom duty",
    COL_RATE_OF_INTEREST_BCD: "RATE OF INTEREST",
    COL_CUST_AIDC: "CUST AIDC",
    COL_RATE_OF_INTEREST_SWS: "RATE OF INTEREST",
    COL_SURCHARGE: "SURCHARGE",
    COL_LAND_COST_WITH_GST: "LAND COST OF PURCHASE WITH GST",
    COL_RATE_PER_UNIT: "Ratepurchase per unitPcs/KGS/SET",
}


def _raw_cell_value(rv: RawValue | None) -> object | None:
    """Resolve the value to write for a directly-extracted ``RawValue`` cell.

    - ``None`` / missing -> ``None`` (the cell is left blank; Req 4.9, 5.11).
    - Unparseable (located but not resolvable) -> the raw text (Req 9.2): the
      printed characters are preserved rather than dropped.
    - Otherwise the parsed interpretation is written when present (full numeric
      precision / verbatim string; Req 8.4), falling back to the raw text.
    """
    if rv is None or rv.is_missing:
        return None
    if rv.is_unparseable:
        return rv.raw_text
    if rv.parsed is not None:
        return rv.parsed
    return rv.raw_text


def _raw_number(rv: RawValue | None) -> float | None:
    """Return the numeric interpretation of a ``RawValue`` (or ``None``).

    Used to total per-line columns whose sums are not pre-computed on ``Totals``
    (e.g. column U ``CUST AIDC`` = BCD amount). Mirrors the column's own written
    value: missing/unparseable/non-numeric yields ``None`` (contributes nothing
    to the sum), preserving full precision otherwise.
    """
    if rv is None or rv.is_missing or rv.is_unparseable:
        return None
    parsed = rv.parsed
    if isinstance(parsed, bool):
        return None
    if isinstance(parsed, (int, float)):
        return float(parsed)
    return None


def _sum_optional(values) -> float:
    """Sum ``float | None`` values, skipping ``None`` (blank cells).

    ``None`` per-line values (a required input was missing/non-numeric, Req 6.13)
    do not contribute. An empty/all-``None`` iterable yields ``0.0`` (Req 7.3).
    Kept at full floating-point precision, matching the sample's ``SUM`` cells.
    """
    total = 0.0
    for value in values:
        if value is not None:
            total += value
    return total


def _shift_coord(coordinate: str, shift: int) -> str:
    """Return ``coordinate`` with its row shifted by ``shift`` rows.

    Used to relocate the auxiliary-section labels (and the ``total usd`` value
    cell) when the Totals_Row floats away from the template anchor: down on
    overflow (positive ``shift``), up when unused blank rows were removed
    (negative ``shift``). A zero ``shift`` returns the coordinate unchanged, so
    a 46-item document keeps the sample's exact cell positions (Req 8.2/8.3).
    """
    if shift == 0:
        return coordinate
    row, col = coordinate_to_tuple(coordinate)
    return f"{get_column_letter(col)}{row + shift}"


class ExcelGenerator:
    """Builds the CTN-layout workbook and returns ``.xlsx`` bytes."""

    def __init__(self, use_formulas: bool = True) -> None:
        """Create a generator.

        ``use_formulas`` (default ``True``) makes the computed cells carry live
        Excel formulas mirroring the sample workbook (e.g. ``L=H*K``,
        ``M=L*<usd_rate>``, ``=SUM(...)`` totals) rather than pre-evaluated
        literal values, so the downloaded workbook recalculates in Excel exactly
        like ``1357 ctn llp.xlsx``. Set it to ``False`` to emit full-precision
        literal values instead (used by tests that read the evaluated numbers
        back without an Excel engine to recalculate them).
        """
        self.use_formulas = use_formulas

    def generate(
        self,
        doc: ComputedDocument,
        flags: ReviewFlagSet,
        grand_total_check: GrandTotalVerification | None = None,
    ) -> bytes:
        """Build ``Sheet1`` in the exact CTN layout and return ``.xlsx`` bytes.

        Constructs the workbook in memory (header block + item table) and
        serialises it. The Totals_Row and auxiliary templates are added by
        task 7.2's ``_write_totals_row`` / ``_write_aux_templates``.

        When ``grand_total_check`` is supplied, the Totals_Row cells whose grand
        total mismatches the BOE declared value are red-filled (Req 11.4, 11.5,
        11.7). When it is ``None`` no fills are applied and the output is
        identical to the un-verified workbook.
        """
        wb = self.build_workbook(doc, flags, grand_total_check)
        buffer = BytesIO()
        wb.save(buffer)
        return buffer.getvalue()

    def build_workbook(
        self,
        doc: ComputedDocument,
        flags: ReviewFlagSet,
        grand_total_check: GrandTotalVerification | None = None,
    ) -> Workbook:
        """Assemble the workbook from the bundled CTN style template.

        Loads ``ctn_template.xlsx`` (which carries the sample's exact column
        widths, row heights, merges, fonts, fills, borders and number formats),
        clears its dynamic cells, then writes this document's header values,
        line items and totals into the already-styled cells. The result is
        visually identical to ``1357 ctn llp.xlsx`` while carrying this
        document's data. Exactly one sheet named ``Sheet1`` (Req 8.1).

        When the document has more line items than the template's styled data
        region (rows 13..60), the totals row and the auxiliary sections below it
        are shifted down so they never overlap the item rows; when it has
        fewer, the unused blank rows are removed so the totals sit two rows
        below the last item instead of leaving a bank of empty rows; the aux
        sections move with the totals either way (see ``_fit_layout``).
        """
        wb = _load_template()
        ws = wb["Sheet1"]

        n_items = len(doc.lines)
        totals_row = _totals_row_for(n_items)
        # Signed displacement from the template anchor: negative when blank
        # rows are removed, positive on overflow, zero for exactly 46 items.
        shift = totals_row - _TEMPLATE_TOTALS_ROW
        if shift:
            self._fit_layout(ws, shift, totals_row)

        self._prepare_template(ws, n_items, totals_row)
        self._write_header_block(ws, doc.header, flags)
        self._write_item_table(ws, doc.lines, doc.header.usd_rate, flags)
        self._write_totals_row(ws, doc.totals, doc.lines, totals_row)
        self._apply_grand_total_fills(ws, totals_row, grand_total_check)
        self._write_aux_templates(ws, shift, totals_row)
        self._write_footer_details(ws, doc.header, shift)
        write_identity_props(wb, doc.header)
        return wb

    # ------------------------------------------------------------------
    # Layout fitting (float the totals row to two blanks below the items)
    # ------------------------------------------------------------------
    def _fit_layout(self, ws: Worksheet, shift: int, totals_row: int) -> None:
        """Move the totals row and aux sections by signed ``shift`` rows.

        Positive ``shift`` (more than 46 items) inserts rows before the
        template totals row, pushing it and every auxiliary section down; the
        new continuation rows get data-row styling. Negative ``shift`` (fewer
        than 46 items) deletes the unused blank data rows starting at the new
        totals position, pulling the template totals row and aux sections up
        so no bank of empty rows remains. openpyxl relocates neither merges
        nor (on delete) reliably all dimensions, so the aux merges at/below
        the template totals row are re-created at their shifted positions.
        """
        # Capture and release the merges that sit at/below the totals row
        # before moving rows. They must be unmerged while their cells still
        # exist (openpyxl does not relocate merge coordinates on
        # insert/delete, so unmerging by the old coordinates afterwards would
        # fail once the cells have shifted).
        below_specs = [
            (rng.min_row, rng.min_col, rng.max_row, rng.max_col)
            for rng in list(ws.merged_cells.ranges)
            if rng.min_row >= _TEMPLATE_TOTALS_ROW
        ]
        for min_row, min_col, max_row, max_col in below_specs:
            ws.unmerge_cells(
                start_row=min_row, start_column=min_col,
                end_row=max_row, end_column=max_col,
            )

        if shift > 0:
            ws.insert_rows(_TEMPLATE_TOTALS_ROW, shift)
        else:
            # Drop data-region scaffolding merges overlapping the removed
            # band (e.g. the sample's G36:G57): their rows disappear, so the
            # merge must not survive shifted onto the aux sections. The item
            # writer needs unmerged cells anyway (see ``_prepare_template``).
            for rng in list(ws.merged_cells.ranges):
                if rng.max_row >= totals_row and rng.min_row < _TEMPLATE_TOTALS_ROW:
                    ws.unmerge_cells(str(rng))
            # Remove the unused blank data rows starting at the new totals
            # position; the template totals row slides up into place.
            ws.delete_rows(totals_row, -shift)

        # Re-create the captured merges at their shifted positions.
        for min_row, min_col, max_row, max_col in below_specs:
            ws.merge_cells(
                start_row=min_row + shift,
                start_column=min_col,
                end_row=max_row + shift,
                end_column=max_col,
            )

        if shift > 0:
            # Style the newly inserted continuation rows like a real data row.
            self._clone_data_row_style(
                ws, _TEMPLATE_TOTALS_ROW, _TEMPLATE_TOTALS_ROW + shift - 1
            )

    # Backwards-compatible alias (overflow-only callers/tests).
    def _expand_for_overflow(self, ws: Worksheet, shift: int) -> None:
        """Insert ``shift`` rows before the template totals row (see ``_fit_layout``)."""
        self._fit_layout(ws, shift, _TEMPLATE_TOTALS_ROW + shift)

    @staticmethod
    def _clone_data_row_style(ws: Worksheet, first_row: int, last_row: int) -> None:
        """Copy the data-row styling onto rows ``first_row..last_row`` (inclusive).

        Clones the per-column style (borders, fonts, fills, number formats) and
        the row height from ``_TEMPLATE_STYLE_ROW`` so overflow item rows render
        identically to the template's styled data region. Values are left blank;
        the item writer fills them in afterwards.
        """
        ref_height = ws.row_dimensions[_TEMPLATE_STYLE_ROW].height
        for row in range(first_row, last_row + 1):
            if ref_height is not None:
                ws.row_dimensions[row].height = ref_height
            for col in range(1, COL_RATE_PER_UNIT + 1):
                src = ws.cell(row=_TEMPLATE_STYLE_ROW, column=col)
                dst = ws.cell(row=row, column=col)
                dst._style = copy(src._style)
                dst.value = None

    # ------------------------------------------------------------------
    # Template preparation (clear the sample's data, keep its styling)
    # ------------------------------------------------------------------
    def _prepare_template(self, ws: Worksheet, n_items: int, totals_row: int) -> None:
        """Clear the template's document-specific cells, preserving all styling.

        Removes the sample's baked-in header values, line items and totals so
        only this document's data is written, while leaving every style, width,
        height and the auxiliary scaffolding intact. Merged ranges that fall in
        the line-item write region are unmerged so each item keeps its own row.
        ``totals_row`` is the floating totals position (two blanks below the
        last item); the layout fit already removed/inserted rows, so rows
        13..totals_row are exactly the items, the two blanks and the totals.
        """
        last_data_row = _TEMPLATE_FIRST_DATA_ROW + max(n_items, 0) - 1
        write_end = max(last_data_row, totals_row)

        # Unmerge any merged range intersecting the data write region (e.g. the
        # sample's G36:G57) so per-row writes never target a merged cell.
        for rng in list(ws.merged_cells.ranges):
            if not (rng.max_row < _TEMPLATE_FIRST_DATA_ROW or rng.min_row > write_end):
                ws.unmerge_cells(str(rng))

        # Clear the header block value cells (labels are rewritten by the header
        # writer); no-source value cells simply stay blank.
        for coord in (
            "E1", "E2", "G2", "E3", "G3", "E4", "G4",
            "E5", "G5", "E6", "G6", "E7", "G7", "E8", "G8",
        ):
            ws[coord] = None

        # Clear the data region and the totals row (cols A..Y), keeping styles.
        for row in range(_TEMPLATE_FIRST_DATA_ROW, write_end + 1):
            for col in range(1, COL_RATE_PER_UNIT + 1):
                ws.cell(row=row, column=col).value = None

        # Harmonize the item-row styling (see ``_normalize_data_region_style``).
        self._normalize_data_region_style(ws, totals_row)

    @staticmethod
    def _normalize_data_region_style(ws: Worksheet, totals_row: int | None = None) -> None:
        """Clone the canonical data-row style across the item rows.

        Copies the per-column style (fonts, fills, borders, number formats) from
        ``_TEMPLATE_STYLE_ROW`` onto every item/blank row below the header so
        the hand-edited inconsistencies in the sample's last rows (e.g. row
        60's stray bold, rows 58-59's missing column-O highlight) do not
        surface. Values are untouched. When ``totals_row`` is given only rows
        above it are touched, so the aux sections below the totals keep their
        own styling; otherwise the template region 13..60 is covered.
        """
        if totals_row is None:
            end_row = _TEMPLATE_LAST_DATA_ROW
        else:
            end_row = min(_TEMPLATE_LAST_DATA_ROW, totals_row - 1)
        for row in range(_TEMPLATE_FIRST_DATA_ROW, end_row + 1):
            if row == _TEMPLATE_STYLE_ROW:
                continue
            for col in range(1, COL_RATE_PER_UNIT + 1):
                src = ws.cell(row=_TEMPLATE_STYLE_ROW, column=col)
                ws.cell(row=row, column=col)._style = copy(src._style)

    # ------------------------------------------------------------------
    # Header block (D1:G8)
    # ------------------------------------------------------------------
    def _write_header_block(
        self, ws: Worksheet, header: HeaderBlock, flags: ReviewFlagSet
    ) -> None:
        """Write the Header_Block labels and values into ``D1:G8`` (Req 4.x).

        Labels (column D/F) are reproduced verbatim from the sample. Values
        (column E/G) are written exactly as extracted, without reformatting the
        number/date values (Req 4.2-4.5). A value flagged missing/unreadable, and
        every no-source field (``Eway bill no/date``, ``RETTENCE DATE/RATE``) is
        left blank with no characters or placeholder (Req 4.8, 4.9, 8.6). The
        ``USD Amt`` (G3) and ``USD Rate`` (G2) cells are now sourced from the BOE
        (Req 12): G3 from the USD invoice amount, G2 from the BOE exchange rate
        with a fallback to the User-supplied rate.
        """
        # --- fixed labels (column D and F) ---
        ws["D1"] = "Company name"
        ws["D2"] = "Party Name"
        ws["F2"] = "USD Rate"
        ws["D3"] = "Details"
        ws["F3"] = "USD Amt"          # value (G3) sourced from the BOE (Req 12.2)
        ws["D4"] = "Invoice No"
        ws["F4"] = "Inv Date"
        ws["D5"] = "BE No"
        ws["F5"] = "BE Date"
        ws["D6"] = "B/L NO"
        ws["F6"] = "B/L DATE"
        ws["D7"] = "Eway bill no"     # no-source label; value cell stays blank
        ws["F7"] = "Eway bill date"   # no-source label; value cell stays blank
        ws["D8"] = "RETTENCE DATE"    # no-source label; value cell stays blank
        ws["F8"] = "RETTENCE RATE"    # no-source label; value cell stays blank

        # --- values (column E and G) ---
        # Company name is a configuration string (Req 4.7), always present.
        if header.company_name:
            ws["E1"] = header.company_name

        self._set_if_value(ws, "E2", _raw_cell_value(header.party_name))   # Req 4.1
        # USD Rate (G2): prefer the BOE-printed exchange rate at full precision
        # (Req 12.3); fall back to the User-supplied rate when the BOE does not
        # print one (Req 12.4). A missing BOE rate is never inferred.
        boe_rate = _raw_cell_value(getattr(header, "usd_rate_boe", None))
        if boe_rate is not None:
            ws["G2"] = boe_rate
        elif header.usd_rate is not None:
            ws["G2"] = header.usd_rate
        # USD Amt (G3): the BOE total invoice amount in USD, full precision, no
        # reformatting; left blank when missing/unresolved (Req 12.2, 12.5).
        self._set_if_value(ws, "G3", _raw_cell_value(header.invoice_amount))  # Req 12.2
        self._set_if_value(ws, "E3", _raw_cell_value(header.details))      # Req 4.6
        self._set_if_value(ws, "E4", _raw_cell_value(header.invoice_no))   # Req 4.2
        self._set_if_value(ws, "G4", _raw_cell_value(header.invoice_date)) # Req 4.2/13.1
        self._set_if_value(ws, "E5", _raw_cell_value(header.be_no))        # Req 4.3
        self._set_if_value(ws, "G5", _raw_cell_value(header.be_date))      # Req 4.3/13.1
        self._set_if_value(ws, "E6", _raw_cell_value(header.bl_no))        # Req 4.4
        self._set_if_value(ws, "G6", _raw_cell_value(header.bl_date))      # Req 4.4/13.1

    # ------------------------------------------------------------------
    # Item table (header row 12, data from row 13)
    # ------------------------------------------------------------------
    def _write_item_table(
        self, ws: Worksheet, lines: list[ComputedLine], usd_rate: float, flags: ReviewFlagSet
    ) -> None:
        """Write the Item_Table header (row 12) and the dense data rows (Req 5.x).

        The header labels are reproduced character-for-character (Req 5.12). Line
        items are written in ascending item-serial order beginning at row 13, one
        per row with no blank rows (Req 5.1); ``Sr. no.`` runs 1..N (Req 5.2).
        Direct values are written verbatim and computed values at full precision
        into their mapped columns (Req 5.3-5.10, 6.11, 8.4, 8.5). Missing/
        flagged cells and no-source columns stay blank (Req 5.11, 8.6); located
        but unparseable fields are written as raw text (Req 9.2).
        """
        # 'In USD' sub-header sits above the 'Amount' column (sample L11).
        ws.cell(row=ITEM_TABLE_HEADER_ROW - 1, column=COL_AMOUNT, value="In USD")

        # Header row 12 - verbatim labels.
        for col, label in ITEM_TABLE_HEADERS.items():
            ws.cell(row=ITEM_TABLE_HEADER_ROW, column=col, value=label)

        # Data rows, ascending serial, dense (no gaps), Sr. no. 1..N.
        ordered = sorted(lines, key=lambda ln: ln.source.item_serial)
        for offset, line in enumerate(ordered):
            row = ITEM_TABLE_FIRST_DATA_ROW + offset
            self._write_item_row(ws, row, offset + 1, line, usd_rate, flags)

    def _write_item_row(
        self,
        ws: Worksheet,
        row: int,
        sr_no: int,
        line: ComputedLine,
        usd_rate: float,
        flags: ReviewFlagSet,
    ) -> None:
        """Write one Line_Item's cells into ``row`` per the column map.

        Columns B (PARTY NAME), C (BILLING AMOUNT), D (AS PER TALLY NAME) and
        G (CTN) have no BOE/configuration source in Milestone 1 and are left
        blank for every row (Req 8.6, mapping table). Review-flagged computed
        values are already ``None`` on the ``ComputedLine`` and so write blank.
        A row whose BOE values differ from the supplier invoice (see
        ``LineItem.invoice_mismatch``) is red-filled across A-Y so the operator
        spots it on open; every other row keeps the template styling.
        """
        item = line.source

        # A: Sr. no. - dense sequential 1..N (Req 5.2).
        ws.cell(row=row, column=COL_SR_NO, value=sr_no)

        # Direct (verbatim) values.
        self._set_cell(ws, row, COL_DESCRIPTION, _raw_cell_value(item.description))
        self._set_cell(ws, row, COL_HSN_CODE, _raw_cell_value(item.cth_hsn))
        # G: per-line CTN - populated from the supplier invoice when supplied,
        # otherwise blank (the BOE itself has no per-line carton count).
        self._set_cell(ws, row, COL_CTN, _raw_cell_value(item.cartons))
        # H/I/K: a line booked under a different quantity than the BOE prints
        # (MTS -> KGS, a weighed tape line -> its piece count; decided by the
        # calculator) writes that quantity and unit, so the workbook matches the
        # Tally stock unit. Amounts are unchanged, so K (UPI) becomes
        # amount_usd / qty to keep L = H*K true; it is blank when the amount is
        # missing (Req 6.13). Every other line is written verbatim.
        if line.stock_qty is not None:
            self._set_cell(ws, row, COL_QTY, line.stock_qty)
            self._set_cell(ws, row, COL_UNIT, line.stock_unit)
            unit_price = (
                line.amount_usd / line.stock_qty
                if line.amount_usd is not None and line.stock_qty
                else None
            )
            self._set_cell(ws, row, COL_UNIT_PRICE_USD, unit_price)
        else:
            self._set_cell(ws, row, COL_QTY, _raw_cell_value(item.quantity))
            self._set_cell(ws, row, COL_UNIT, _raw_cell_value(item.unit))
            self._set_cell(ws, row, COL_UNIT_PRICE_USD, _raw_cell_value(item.unit_price_usd))
        self._set_cell(ws, row, COL_CUSTOM_ASS_VALUE, _raw_cell_value(item.assessable_value))
        # Direct (verbatim) rate values. BOE rates are stored as fractions
        # (e.g. 0.075 for 7.5%). The template's whole-percent format rounds
        # fractional rates, so R/T get an explicit value-aware percent format.
        self._set_rate_cell(ws, row, COL_RATE_OF_DUTY_IGST, _raw_cell_value(item.igst_rate))
        self._set_rate_cell(ws, row, COL_RATE_OF_INTEREST_BCD, _raw_cell_value(item.bcd_rate))
        # U: CUST AIDC = BCD amount + CHCESS (computed); falls back to the raw
        # BCD amount when no cess is present (cust_aidc == bcd_amount).
        self._set_cell(ws, row, COL_CUST_AIDC, line.cust_aidc)

        # SWS rate (column V): write the ACTUAL per-line rate the BOE declares
        # when it is available (an exemption 0 is honoured as 0); otherwise keep
        # the historical 0.1 default so existing behaviour is preserved (Req 15).
        sws_rate_val = _raw_number(item.sws_rate)
        if sws_rate_val is not None:
            self._set_rate_cell(ws, row, COL_RATE_OF_INTEREST_SWS, sws_rate_val)
        else:
            ws.cell(row=row, column=COL_RATE_OF_INTEREST_SWS, value=SWS_RATE)

        # GST/IGST amount (column Q) is a literal in the sample (manually keyed),
        # so it is always written as a value and is referenced by the S and X
        # formulas below.
        self._set_cell(ws, row, COL_GST, line.igst_amount)

        # Computed (full-precision) values. When formula mode is on these carry
        # the sample's live Excel formulas (referencing the literal cells H, K,
        # U, V, Q on the same row); otherwise the pre-evaluated literal is
        # written. A ``None`` computed value (missing/non-numeric input, Req
        # 6.13) leaves the cell blank in both modes so the formula never turns a
        # missing input into a spurious 0.
        qty = line.stock_qty if line.stock_qty is not None else _raw_number(item.quantity)
        # pcs = qty * pcs_factor(unit); the factor comes from the shared units
        # table via the calculator (DOZ=12, GRS=144, THD=1000). Render it without
        # a trailing .0 for whole numbers (via _num) so the formula reads
        # ``=H13*144`` not ``=H13*144.0``. When the line has no factor, line.pcs
        # is None and _put_computed leaves the cell blank.
        pcs_formula = (
            f"=H{row}*{self._num(line.pcs_factor)}"
            if line.pcs_factor is not None
            else f"=H{row}"
        )
        self._put_computed(ws, row, COL_PCS, line.pcs, pcs_formula)
        self._put_computed(ws, row, COL_AMOUNT, line.amount_usd, f"=H{row}*K{row}")
        self._put_computed(
            ws, row, COL_RATE_PER_USD, line.purchase_inr, f"=L{row}*{self._num(usd_rate)}"
        )
        self._put_computed(ws, row, COL_SURCHARGE, line.sws_amount, f"=U{row}*V{row}")
        self._put_computed(
            ws, row, COL_TOTAL_CUSTOM_DUTY, line.total_customs_duty, f"=U{row}+W{row}"
        )
        self._put_computed(
            ws, row, COL_LAND_COST_WITHOUT_GST, line.land_cost_excl_gst, f"=M{row}+P{row}"
        )
        self._put_computed(
            ws, row, COL_TOTAL_CUSTOM_DUTY_2, line.combined_duty, f"=P{row}+Q{row}"
        )
        self._put_computed(
            ws, row, COL_LAND_COST_WITH_GST, line.land_cost_incl_gst, f"=M{row}+S{row}"
        )
        # Y = O/H; guard the qty==0 case (calculator yields 0) so the formula
        # never produces a #DIV/0! error in the workbook. ``qty`` is what H
        # holds (the booking quantity), which is also the calculator's divisor,
        # so the formula and the literal agree.
        if line.purchase_rate_per_unit is not None and self.use_formulas and qty not in (None, 0):
            ws.cell(row=row, column=COL_RATE_PER_UNIT, value=f"=O{row}/H{row}")
        else:
            self._set_cell(ws, row, COL_RATE_PER_UNIT, line.purchase_rate_per_unit)

        # BOE-vs-invoice mismatch: red-fill the whole item row (A-Y) when the
        # orchestrator found this serial's quantity, unit price or amount
        # differing between the two PDFs. Applied last so it only restyles -
        # values, formulas, borders and number formats are untouched.
        if item.invoice_mismatch:
            for col in range(COL_SR_NO, COL_RATE_PER_UNIT + 1):
                ws.cell(row=row, column=col).fill = INVOICE_MISMATCH_FILL

    # ------------------------------------------------------------------
    # Totals row (two blank rows below the last item)
    # ------------------------------------------------------------------
    def _write_totals_row(
        self, ws: Worksheet, totals: Totals, lines: list[ComputedLine], totals_row: int
    ) -> None:
        """Write the Totals_Row at ``totals_row`` (two blanks below the items).

        Reproduces the sample's summed columns (``G, L, M, N, O, P, Q, S, U, W,
        X``) at full floating-point precision (Req 7.4, 8.2). Columns backed by
        a pre-computed ``Totals`` field use it directly (L, N, O, P, Q, X); the
        remaining column sums (G, M, S, U, W) are summed here from the per-line
        values exactly as written into those columns, matching the sample's
        ``SUM`` formulas. Column G carries the ``SUM`` of the per-line carton
        cells when the invoice supplied any carton counts, else the BOE total
        package/CTN count taken from ``Totals.package_count`` (Req 7.5); if
        that is missing the cell is left blank rather than substituting a value.
        """
        row = totals_row
        n = len(lines)
        first = ITEM_TABLE_FIRST_DATA_ROW
        last = first + n - 1  # inclusive last data row; < first when no items

        # G: CTN total = SUM of the per-line carton cells when the invoice
        # supplied any carton counts (a live =SUM formula in formula mode, the
        # pre-evaluated sum otherwise), so the total always agrees with the
        # column above it. With no per-line cartons (no invoice supplied) it
        # falls back to the BOE total package count (Req 7.5), blank when that
        # is missing rather than substituting a value.
        carton_vals = [
            v for v in (_raw_number(ln.source.cartons) for ln in lines)
            if v is not None
        ]
        if carton_vals:
            self._put_total(ws, row, COL_CTN, "G", first, last, sum(carton_vals))
        else:
            self._set_cell(ws, row, COL_CTN, _raw_cell_value(totals.package_count))

        # The remaining summed columns mirror the sample's ``=SUM(<col>13:<col>N)``
        # formulas in formula mode, or the pre-evaluated literal otherwise.
        self._put_total(ws, row, COL_AMOUNT, "L", first, last, totals.total_amount_usd)
        self._put_total(
            ws, row, COL_RATE_PER_USD, "M", first, last,
            _sum_optional(ln.purchase_inr for ln in lines),
        )
        self._put_total(
            ws, row, COL_CUSTOM_ASS_VALUE, "N", first, last, totals.total_assessable_value
        )
        self._put_total(
            ws, row, COL_LAND_COST_WITHOUT_GST, "O", first, last,
            totals.total_land_cost_excl_gst,
        )
        self._put_total(
            ws, row, COL_TOTAL_CUSTOM_DUTY, "P", first, last, totals.total_customs_duty
        )
        self._put_total(ws, row, COL_GST, "Q", first, last, totals.total_igst)
        self._put_total(
            ws, row, COL_TOTAL_CUSTOM_DUTY_2, "S", first, last,
            _sum_optional(ln.combined_duty for ln in lines),
        )
        self._put_total(
            ws, row, COL_CUST_AIDC, "U", first, last,
            _sum_optional(ln.cust_aidc for ln in lines),
        )
        self._put_total(
            ws, row, COL_SURCHARGE, "W", first, last,
            _sum_optional(ln.sws_amount for ln in lines),
        )
        self._put_total(
            ws, row, COL_LAND_COST_WITH_GST, "X", first, last,
            totals.total_land_cost_incl_gst,
        )

    def _apply_grand_total_fills(
        self,
        ws: Worksheet,
        totals_row: int,
        grand_total_check: GrandTotalVerification | None,
    ) -> None:
        """Red-fill the Totals_Row cells whose grand total mismatches the BOE.

        For each column letter in ``grand_total_check.highlight_columns`` (the
        MISMATCH columns; Req 11.4) the mapped Totals_Row cell (P -> TOTAL Custom
        Duty, Q -> GST, S -> total custom duty) is given a solid red fill (Req
        11.5, 11.7). MATCH and UNVERIFIED columns are left unstyled. When
        ``grand_total_check`` is ``None`` nothing is changed, so the un-verified
        output is preserved exactly.
        """
        if grand_total_check is None:
            return
        for column in grand_total_check.highlight_columns:
            col_index = GRAND_TOTAL_FILL_COLUMNS.get(column)
            if col_index is None:
                continue
            ws.cell(row=totals_row, column=col_index).fill = GRAND_TOTAL_MISMATCH_FILL

    # ------------------------------------------------------------------
    # Auxiliary template sections (labels only; data cells empty)
    # ------------------------------------------------------------------
    def _write_aux_templates(self, ws: Worksheet, shift: int, totals_row: int) -> None:
        """Write the auxiliary section labels verbatim at their (shifted) positions.

        Reproduces character-for-character (Req 8.3): the ``total usd`` summary
        (B71 label + E71 value mirroring ``=L<totals>``), the ``DETAILS AS PER
        CHALLANS`` and ``DETAILS AS PER TALLY`` section titles with their
        column-header rows, the C&F detail block row labels (B72-B86), and the
        ``CLEARANCE AND FORWARDING INVOICE`` title with its header row. When the
        line-item count overflows the template's data region these positions are
        shifted down by ``shift`` rows so they sit below the item rows rather
        than colliding with them; with no overflow (``shift == 0``) they remain
        at the sample's exact cell positions (Req 8.2). The data cells within
        these sections are intentionally left empty for Milestone 1.
        """
        for coordinate, label in AUX_LABELS.items():
            ws[_shift_coord(coordinate, shift)] = label

        # 'total usd' value cell: the sample stores '=L61'; mirror that with a
        # live formula pointing at the floating totals row in formula mode,
        # Totals_Row L cell holds the identical value/formula.
        target = _shift_coord(AUX_TOTAL_USD_CELL, shift)
        if self.use_formulas:
            ws[target] = f"=L{totals_row}"
        else:
            total_usd = ws.cell(row=totals_row, column=COL_AMOUNT).value
            if total_usd is not None:
                ws[target] = total_usd

    # ------------------------------------------------------------------
    # 'customer and other expenses' footer detail cells (column E)
    # ------------------------------------------------------------------
    def _write_footer_details(
        self, ws: Worksheet, header: HeaderBlock, shift: int
    ) -> None:
        """Populate the 'customer and other expenses' footer (column E) from THIS
        document's header.

        The workbook is built by cloning a style template that carries the
        original sample BOE's values in these cells (E75 B/E NO, E76 B/E DATE,
        E77 BL NO, E78 BL DATE). Without this step those template values bleed
        through and appear "defaulted" on every conversion. Each cell is always
        overwritten - with the BOE value, or cleared (``None``) when the field is
        missing - so no stale template data survives (mirrors the top Header_Block,
        Req 4.3/4.4). Rows shift down by ``shift`` under item-table overflow.
        """
        footer_col = 5  # column E - the value column of the footer block
        for base_row, raw in (
            (75, header.be_no),    # B/E NO
            (76, header.be_date),  # B/E DATE
            (77, header.bl_no),    # BL NO
            (78, header.bl_date),  # BL DATE
        ):
            ws.cell(row=base_row + shift, column=footer_col).value = _raw_cell_value(raw)

    # ------------------------------------------------------------------
    # Small writing helpers
    # ------------------------------------------------------------------
    def _put_computed(
        self, ws: Worksheet, row: int, col: int, value: object | None, formula: str
    ) -> None:
        """Write a computed cell as a formula or literal, blank when ``None``.

        In formula mode the live Excel ``formula`` is written (mirroring the
        sample workbook); otherwise the pre-evaluated ``value`` literal is
        written. A ``None`` value (missing/non-numeric input, Req 6.13) leaves
        the cell blank in both modes, so a formula is never written for a cell
        whose inputs were absent (which would otherwise read blanks as 0).
        """
        if value is None:
            return
        if self.use_formulas:
            ws.cell(row=row, column=col, value=formula)
        else:
            ws.cell(row=row, column=col, value=value)

    def _put_total(
        self,
        ws: Worksheet,
        row: int,
        col: int,
        col_letter: str,
        first: int,
        last: int,
        value: object | None,
    ) -> None:
        """Write a Totals_Row cell as ``=SUM(<col>first:<col>last)`` or a literal.

        In formula mode a ``SUM`` over the actual data rows is written (matching
        the sample's ``=SUM(L13:L59)`` style); with no data rows
        (``last < first``) or in literal mode the pre-evaluated ``value`` is
        written instead.
        """
        if self.use_formulas and last >= first:
            ws.cell(row=row, column=col, value=f"=SUM({col_letter}{first}:{col_letter}{last})")
        else:
            self._set_cell(ws, row, col, value)

    @staticmethod
    def _num(value: float) -> str:
        """Render a numeric literal for embedding in a formula (no exponent).

        Integers render without a trailing ``.0`` (e.g. ``95`` not ``95.0``) and
        finite floats render in plain decimal form so the formula text matches
        the sample's style (``=L13*95.3``).
        """
        if value is None:
            return "0"
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return repr(value)

    @staticmethod
    def _set_cell(ws: Worksheet, row: int, col: int, value: object | None) -> None:
        """Write ``value`` at ``(row, col)`` only when it is not ``None``.

        Leaving ``None`` unwritten keeps no-source/flagged cells truly empty (no
        characters, spaces, or placeholder text) per Req 4.8/4.9/5.11/8.6.
        """
        if value is not None:
            ws.cell(row=row, column=col, value=value)

    @staticmethod
    def _set_rate_cell(ws: Worksheet, row: int, col: int, value: object | None) -> None:
        """Write a direct rate value with whole/fractional percent display.

        Whole percentages use ``0%`` so 18% renders without a dangling decimal
        point. Fractional percentages use ``0.##%`` so 7.5% is not rounded to 8%.
        """
        if value is None:
            return
        cell = ws.cell(row=row, column=col, value=value)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            percent = float(value) * 100
            cell.number_format = "0%" if abs(percent - round(percent)) < 1e-9 else "0.##%"

    @staticmethod
    def _set_if_value(ws: Worksheet, coordinate: str, value: object | None) -> None:
        """Write ``value`` at ``coordinate`` only when it is not ``None``."""
        if value is not None:
            ws[coordinate] = value


# ---------------------------------------------------------------------------
# Buyer / seller identity carried through the Excel round-trip
# ---------------------------------------------------------------------------
# The CTN sheet is a costing table: it has no GSTIN, state or country cells, and
# adding some would change a layout that is reproduced character-for-character
# from the sample workbook (and golden-tested). But the *upload an edited Excel*
# path needs that identity - without it the Tally voucher goes out with a blank
# ``countryofresidence`` / ``placeofsupply`` / ``consigneestatename``, and Tally
# fills a blank field with its own company defaults (showing India for a Chinese
# supplier). So the identity travels as OOXML custom document properties:
# invisible in the grid, outside every golden-tested cell, unaffected by the
# one-sheet rule (Req 8.1), and preserved when Excel re-saves the file.
IDENTITY_PROP_PREFIX = "boe."

# HeaderBlock field name -> document property name.
IDENTITY_FIELDS: tuple[str, ...] = (
    "buyer_gstin",
    "buyer_state",
    "buyer_pincode",
    "buyer_address",
    "seller_country",
    "seller_address",
)


def _identity_text(rv: RawValue | None) -> str:
    """A RawValue as plain text for a document property ("" when absent)."""
    if rv is None or rv.is_missing:
        return ""
    parsed = getattr(rv, "parsed", None)
    if isinstance(parsed, str) and parsed.strip():
        return parsed.strip()
    raw = getattr(rv, "raw_text", None)
    return raw.strip() if isinstance(raw, str) and raw.strip() else ""


def write_identity_props(wb: Workbook, header: HeaderBlock) -> None:
    """Attach the buyer/seller identity to ``wb`` as custom document properties.

    Only fields the document actually carries are written, so a BOE that never
    stated a value does not gain an empty property.
    """
    from openpyxl.packaging.custom import StringProperty

    for field_name in IDENTITY_FIELDS:
        value = _identity_text(getattr(header, field_name, None))
        if not value:
            continue
        wb.custom_doc_props.append(
            StringProperty(name=f"{IDENTITY_PROP_PREFIX}{field_name}", value=value)
        )


def read_identity_props(wb: Workbook) -> dict[str, str]:
    """The identity fields stored on ``wb`` (empty for a pre-change workbook)."""
    props = getattr(wb, "custom_doc_props", None)
    if props is None:
        return {}
    out: dict[str, str] = {}
    for prop in props:
        name = getattr(prop, "name", "") or ""
        if not name.startswith(IDENTITY_PROP_PREFIX):
            continue
        field_name = name[len(IDENTITY_PROP_PREFIX) :]
        value = getattr(prop, "value", None)
        if field_name in IDENTITY_FIELDS and isinstance(value, str) and value.strip():
            out[field_name] = value.strip()
    return out
