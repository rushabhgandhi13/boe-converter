"""PDF_Parser: orientation-aware extraction of BOE header and line items.

This module implements the orientation-aware extraction *primitives* used by the
parser (task 4.1):

- :meth:`PdfParser._upright_words` keeps only characters whose orientation is 0
  (i.e. normal, horizontal reading direction) and reconstructs words from them.
  This defeats the BOE's rotated margin labels (e.g. ``SLIATED``/``YTUD``/
  ``SEITUD``) which naive full-page text extraction would otherwise interleave
  with the real horizontal data.
- :meth:`PdfParser._reconstruct_rows` groups upright words into geometric rows by
  their vertical bounding box so downstream stages can read tabular data
  positionally.
- :meth:`PdfParser._capture` is the verbatim-capture helper: it records a field
  as a :class:`RawValue` preserving the printed ``raw_text`` and performs a
  separate, non-destructive numeric parse, setting ``is_missing``/
  ``is_unparseable`` as appropriate (Req 2.11, 3.6, 9.2, 9.3).

The header (task 4.4) and line-item (tasks 5.x) extraction stages build on these
primitives and are implemented separately.

Design references: design.md "PDF_Parser" and "Data Models -> RawValue".
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass

import pdfplumber
from pdfplumber.utils import extract_words

from boe_converter.models import (
    ExtractedDocument,
    HeaderBlock,
    LineItem,
    RawValue,
    ReviewFlag,
)

# A date as printed in the BOE: either ``07-JUN-26`` (Part II invoice date) or
# ``22/06/2026`` (Part I banded dates). Used to locate date *values* positionally
# when a column carries both an identifier and a date in stacked rows.
_DATE_RE = re.compile(
    r"\d{1,2}[-/][A-Za-z]{3,9}[-/]\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4}"
)


@dataclass(frozen=True)
class Word:
    """A single upright word with its bounding box.

    Coordinates follow pdfplumber's top-origin convention: ``top``/``bottom`` are
    measured from the top of the page, ``x0``/``x1`` from the left edge.
    """

    text: str
    x0: float
    x1: float
    top: float
    bottom: float

    @property
    def mid_y(self) -> float:
        """Vertical midpoint of the word's bounding box."""
        return (self.top + self.bottom) / 2.0


@dataclass(frozen=True)
class InvoiceItemRow:
    """Part II (invoice & valuation) values for a single line item.

    Captured verbatim from the BOE's ``1.S NO. | 2.CTH | 3.DESCRIPTION |
    4.UNIT PRICE | 5.QUANTITY | 6.UQC | 7.AMOUNT`` table (Req 3.5-3.8). The
    intermediate row is joined with its :class:`DutyItemRow` counterpart on
    ``item_serial`` by the merge stage (task 5.2) to build a ``LineItem``.
    """

    item_serial: int            # 1.S NO. (Req 3.4, join key)
    cth_hsn: RawValue           # 2.CTH (Req 3.5)
    description: RawValue       # 3.DESCRIPTION, verbatim incl. wrapped lines (Req 3.6)
    unit_price_usd: RawValue    # 4.UNIT PRICE / UPI (Req 3.7)
    quantity: RawValue          # 5.QUANTITY (Req 3.8)
    unit: RawValue              # 6.UQC (Req 3.8)
    amount: RawValue            # 7.AMOUNT (invoice line amount, for cross-check)


@dataclass(frozen=True)
class DutyItemRow:
    """Part III (duties) values for a single line item.

    Captured verbatim from the BOE's per-item duty block keyed by ``2.ITEMSN``
    (Req 3.9-3.12). ``bcd_rate``/``igst_rate`` preserve the printed percentage in
    ``raw_text`` while ``parsed`` holds the decimal-fraction interpretation
    (e.g. printed ``"18"`` -> ``parsed == 0.18``) consumed by the calculator.
    Exemption-driven zero rates/amounts are captured as numeric ``0`` (Req 3.14).
    """

    item_serial: int            # 2.ITEMSN (Req 3.4, join key)
    assessable_value: RawValue  # 29.ASSESS VALUE (Req 3.9)
    bcd_rate: RawValue          # BCD Rate, as a decimal fraction (Req 3.10)
    bcd_amount: RawValue        # BCD Amount (Req 3.10)
    sws_amount: RawValue        # SWS Amount (Req 3.10 surcharge component)
    igst_rate: RawValue         # IGST Rate, as a decimal fraction (Req 3.11)
    total_duty: RawValue        # 30. TOTAL DUTY (Req 3.12)


@dataclass(frozen=True)
class _InvCols:
    """Vertical split lines (x-coordinates) for the Part II invoice table.

    A token's column is decided by where its horizontal centre falls relative to
    these splits. ``serial`` and ``CTH`` are taken positionally as the first two
    tokens left of ``desc_up`` (the description starts further left than its
    header label, so a label-derived left edge cannot bound it).
    """

    desc_up: float    # description | unit-price boundary
    up_qty: float     # unit-price | quantity boundary
    qty_uqc: float    # quantity | UQC boundary
    uqc_amt: float    # UQC | amount boundary


@dataclass(frozen=True)
class _DutyCols:
    """Horizontal value ranges (x-coordinate windows) for a Part III duty grid."""

    bcd: tuple[float, float]
    sws: tuple[float, float]
    igst: tuple[float, float]


def _char_orientation(char: dict) -> int:
    """Return a character's text rotation in degrees, rounded to an int.

    pdfplumber exposes the text matrix ``(a, b, c, d, e, f)``; the writing-
    direction angle is ``atan2(b, a)``. Upright, horizontally-typeset glyphs have
    ``a == 1, b == 0`` => 0 degrees. Rotated margin labels have non-zero angles
    (e.g. ~90 for vertical text, ~180 for upside-down/reversed text).
    """
    matrix = char.get("matrix")
    if not matrix:
        # Fall back to pdfplumber's own upright flag when no matrix is present.
        return 0 if char.get("upright", True) else 90
    a, b = matrix[0], matrix[1]
    return round(math.degrees(math.atan2(b, a)))


class PdfParser:
    """Extracts an :class:`ExtractedDocument` from an opened BOE PDF handle."""

    # Words within this many points of one another (vertically) are treated as
    # belonging to the same geometric row. The BOE body text is ~8-10pt, so a
    # 3pt tolerance reliably groups a single printed line without merging
    # adjacent lines.
    ROW_TOLERANCE = 3.0

    def parse(
        self,
        doc,
        *,
        company_name: str = "Gemini Unicom LLP",
        usd_rate: float = 0.0,
    ) -> ExtractedDocument:
        """Parse a BOE into a fully-populated :class:`ExtractedDocument`.

        ``doc`` may be an already-opened ``pdfplumber.PDF`` handle (the validated
        handle threaded through by the orchestrator) or a filesystem path/bytes
        buffer this method opens itself. The orchestrator-supplied configuration
        (``company_name``, Open Q6) and User-supplied ``usd_rate`` (Open Q5) are
        threaded into the returned ``HeaderBlock``; they are not read from the
        PDF.

        Pipeline (design "PDF_Parser"):

        1. ``_extract_header`` -> Header_Block + header review flags (Req 2.x).
        2. ``_extract_invoice_items`` -> Part II ``{serial: InvoiceItemRow}``.
        3. ``_extract_duty_items`` -> Part III ``{serial: DutyItemRow}``.
        4. ``_extract_declared_item_count`` -> the BOE's declared item count.
        5. ``_merge_items`` -> joins Part II/III on ``item_serial`` into the
           ascending ``list[LineItem]`` plus per-line review flags (Req 3.x/9.x).

        All review flags (header + line item) are aggregated onto the returned
        document. Nothing is silently dropped: a serial present in either source
        yields a ``LineItem`` and any missing/unparseable field is flagged.
        """
        handle, pages, should_close = self._resolve(doc)
        try:
            header, header_flags = self._extract_header(
                pages, company_name=company_name, usd_rate=usd_rate
            )
            invoice_items = self._extract_invoice_items(pages)
            duty_items = self._extract_duty_items(pages)
            declared_count = self._extract_declared_item_count(pages)
            line_items, item_flags = self._merge_items(
                invoice_items, duty_items, declared_count
            )
            return ExtractedDocument(
                header=header,
                line_items=line_items,
                declared_item_count=declared_count,
                flags=[*header_flags, *item_flags],
            )
        finally:
            if should_close and handle is not None:
                try:
                    handle.close()
                except Exception:
                    pass

    @staticmethod
    def _resolve(doc):
        """Return ``(handle, pages, should_close)`` for ``doc``.

        Accepts an already-opened ``pdfplumber.PDF`` (anything exposing
        ``.pages``) and leaves ownership with the caller (``should_close`` is
        ``False``); otherwise treats ``doc`` as a path/bytes buffer, opens it
        with ``pdfplumber`` and takes ownership (``should_close`` is ``True``).
        """
        if hasattr(doc, "pages"):
            return doc, list(doc.pages), False
        handle = pdfplumber.open(doc)
        return handle, list(handle.pages), True

    # ------------------------------------------------------------------
    # Orientation-aware extraction primitives (task 4.1)
    # ------------------------------------------------------------------
    def _upright_words(self, page) -> list[Word]:
        """Return only the upright (orientation == 0) words on ``page``.

        Characters whose text rotation is not 0 degrees -- the rotated/vertical
        margin labels such as ``SLIATED``/``YTUD``/``SEITUD`` -- are discarded
        before word grouping, so rotated text can never contaminate the
        extracted rows (the rotated-label hazard, design.md).

        Words are returned in reading order (top-to-bottom, then left-to-right).
        """
        upright_chars = [ch for ch in page.chars if _char_orientation(ch) == 0]
        if not upright_chars:
            return []

        raw_words = extract_words(
            upright_chars,
            keep_blank_chars=False,
            use_text_flow=False,
        )

        words = [
            Word(
                text=w["text"],
                x0=float(w["x0"]),
                x1=float(w["x1"]),
                top=float(w["top"]),
                bottom=float(w["bottom"]),
            )
            for w in raw_words
        ]
        words.sort(key=lambda w: (round(w.top, 1), w.x0))
        return words

    def _reconstruct_rows(
        self, words: list[Word], tolerance: float | None = None
    ) -> list[list[Word]]:
        """Group ``words`` into geometric rows by their vertical position.

        Words whose vertical midpoints fall within ``tolerance`` points of the
        current row's running average are placed on the same row. Each returned
        row is ordered left-to-right by ``x0``; rows are ordered top-to-bottom.
        This positional (bounding-box) reconstruction is what lets downstream
        stages read the BOE's tables without relying on naive ``extract_text``.
        """
        if tolerance is None:
            tolerance = self.ROW_TOLERANCE
        if not words:
            return []

        ordered = sorted(words, key=lambda w: (w.mid_y, w.x0))
        rows: list[list[Word]] = []
        current: list[Word] = [ordered[0]]
        current_y = ordered[0].mid_y

        for word in ordered[1:]:
            if abs(word.mid_y - current_y) <= tolerance:
                current.append(word)
                # Running mean keeps the reference stable across slight baseline
                # jitter within a single printed line.
                current_y = sum(w.mid_y for w in current) / len(current)
            else:
                current.sort(key=lambda w: w.x0)
                rows.append(current)
                current = [word]
                current_y = word.mid_y

        current.sort(key=lambda w: w.x0)
        rows.append(current)
        return rows

    # ------------------------------------------------------------------
    # Verbatim capture helper (task 4.1)
    # ------------------------------------------------------------------
    @staticmethod
    def _parse_number(raw_text: str) -> float | None:
        """Non-destructively parse a printed value into a float.

        Strips grouping commas, surrounding currency symbols/whitespace and a
        trailing percent sign, then attempts ``float``. Returns ``None`` when the
        text does not represent a single number. The original ``raw_text`` is
        never modified -- this is a read-only interpretation step.
        """
        if raw_text is None:
            return None
        text = raw_text.strip()
        if not text:
            return None

        # Remove common currency symbols and grouping commas; keep sign/decimal.
        cleaned = text.replace(",", "")
        cleaned = re.sub(r"[₹$€£]", "", cleaned)
        cleaned = cleaned.strip()

        percent = cleaned.endswith("%")
        if percent:
            cleaned = cleaned[:-1].strip()

        if not re.fullmatch(r"[+-]?\d*\.?\d+", cleaned):
            return None
        try:
            value = float(cleaned)
        except ValueError:
            return None
        return value / 100.0 if percent else value

    def _capture(self, raw_text: str | None, *, numeric: bool = False) -> RawValue:
        """Capture a field verbatim as a :class:`RawValue`.

        The printed text is preserved exactly in ``raw_text``; a separate,
        non-destructive numeric parse populates ``parsed`` without altering the
        stored characters (Req 2.11, 3.6).

        - ``raw_text is None`` or blank -> :meth:`RawValue.missing` (the field
          could not be located).
        - ``numeric=True`` and the text cannot be parsed as a number ->
          :meth:`RawValue.unparseable`, retaining the raw text (Req 9.2/9.3).
        - ``numeric=False`` -> the raw string is preserved; ``parsed`` holds the
          numeric interpretation when the text happens to be a clean number,
          otherwise the verbatim string itself (text is always resolvable as
          text, so it is never marked unparseable).
        """
        if raw_text is None or raw_text.strip() == "":
            return RawValue.missing()

        parsed_number = self._parse_number(raw_text)

        if numeric:
            if parsed_number is None:
                return RawValue.unparseable(raw_text)
            return RawValue(
                raw_text=raw_text,
                parsed=parsed_number,
                is_missing=False,
                is_unparseable=False,
            )

        return RawValue(
            raw_text=raw_text,
            parsed=parsed_number if parsed_number is not None else raw_text,
            is_missing=False,
            is_unparseable=False,
        )

    # ------------------------------------------------------------------
    # Header-block extraction (task 4.4)
    # ------------------------------------------------------------------
    def _extract_header(
        self,
        pages,
        *,
        company_name: str = "Gemini Unicom LLP",
        usd_rate: float = 0.0,
    ) -> tuple[HeaderBlock, list[ReviewFlag]]:
        """Extract the document-level Header_Block fields from a BOE.

        Uses the orientation-aware upright-word rows (``_upright_words`` +
        ``_reconstruct_rows``) so the BOE's rotated margin labels never
        contaminate the read. Each field is located *positionally* by matching
        its printed column label and then reading the aligned value below/beside
        it, and captured verbatim via :meth:`_capture` (Req 2.11).

        The fields required by Req 2.1-2.8 (BE No/Date, Invoice No/Date, invoice
        amount + currency, package count, party name, container details) each
        emit a ``ReviewFlag(header, field, MISSING/UNPARSEABLE)`` when they cannot
        be located or resolved; no inferred or default value is ever substituted
        (Req 2.10). The Bill of Lading number/date (Req 2.9) are extracted only
        *where present* and are never flagged when absent.

        ``company_name`` (configuration, Open Q6) and ``usd_rate`` (User-supplied,
        Open Q5) are not read from the PDF; they are accepted here so the
        orchestrator can thread them through into the returned ``HeaderBlock``.
        The ``details`` field (Req 4.6, e.g. ``"CO-32 CTN-1357"``) is outside the
        Req 2.x scope of this stage and is left missing here.
        """
        # Reconstruct each page's geometric rows exactly once.
        pages_rows = [
            self._reconstruct_rows(self._upright_words(page)) for page in pages
        ]

        be_no, be_date = self._extract_be_no_date(pages_rows)
        invoice_no, invoice_amount, invoice_currency = self._extract_invoice_summary(
            pages_rows
        )
        invoice_date = self._extract_invoice_date(pages_rows)
        package_count = self._extract_package_count(pages_rows)
        party_name = self._extract_party_name(pages_rows)
        container_details = self._extract_container_details(pages_rows)
        bl_no, bl_date = self._extract_bl_no_date(pages_rows)  # where present

        header = HeaderBlock(
            company_name=company_name,
            party_name=party_name,
            usd_rate=usd_rate,
            details=RawValue.missing(),  # Req 4.6 - out of scope for task 4.4
            invoice_no=invoice_no,
            invoice_date=invoice_date,
            be_no=be_no,
            be_date=be_date,
            bl_no=bl_no,
            bl_date=bl_date,
            invoice_amount=invoice_amount,
            invoice_currency=invoice_currency,
            package_count=package_count,
            container_details=container_details,
        )

        # Required header fields (Req 2.1-2.8). B/L (2.9) is "where present" and
        # is intentionally not in this list.
        required: list[tuple[str, RawValue]] = [
            ("be_no", be_no),
            ("be_date", be_date),
            ("invoice_no", invoice_no),
            ("invoice_date", invoice_date),
            ("invoice_amount", invoice_amount),
            ("invoice_currency", invoice_currency),
            ("package_count", package_count),
            ("party_name", party_name),
            ("container_details", container_details),
        ]
        flags: list[ReviewFlag] = []
        for name, value in required:
            if value.is_missing:
                flags.append(
                    ReviewFlag(scope="header", field_name=name, reason="MISSING")
                )
            elif value.is_unparseable:
                flags.append(
                    ReviewFlag(
                        scope="header",
                        field_name=name,
                        reason="UNPARSEABLE",
                        raw_text=value.raw_text,
                    )
                )

        return header, flags

    # ------------------------------------------------------------------
    # Positional search helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _center(word: Word) -> float:
        """Horizontal midpoint of a word's bounding box."""
        return (word.x0 + word.x1) / 2.0

    def _find_label_row(self, pages_rows, predicate):
        """Return ``(page_index, row_index, row)`` of the first row satisfying
        ``predicate`` over its list of :class:`Word`, scanning pages then rows in
        reading order; ``None`` if no row matches."""
        for pi, rows in enumerate(pages_rows):
            for ri, row in enumerate(rows):
                if predicate(row):
                    return pi, ri, row
        return None

    @staticmethod
    def _word_where(row, predicate) -> int | None:
        """Index of the first word in ``row`` whose text satisfies ``predicate``."""
        for i, w in enumerate(row):
            if predicate(w.text):
                return i
        return None

    def _value_near_anchor(
        self,
        rows,
        label_idx: int,
        anchor_x: float,
        *,
        max_dist: float = 30.0,
        max_scan: int = 4,
        x_range: tuple[float, float] | None = None,
    ) -> str | None:
        """Read a value aligned (horizontally) under ``anchor_x``.

        Scans the rows immediately below ``label_idx`` (up to ``max_scan`` rows)
        and returns the text of the first word whose horizontal centre is closest
        to ``anchor_x`` within ``max_dist`` points (and inside ``x_range`` when
        given). Returns ``None`` when no aligned value is found, so an empty
        column never grabs an unrelated neighbour.
        """
        for ri in range(label_idx + 1, min(label_idx + 1 + max_scan, len(rows))):
            best: Word | None = None
            best_d: float | None = None
            for w in rows[ri]:
                c = self._center(w)
                if x_range is not None and not (x_range[0] <= c <= x_range[1]):
                    continue
                d = abs(c - anchor_x)
                if best_d is None or d < best_d:
                    best_d, best = d, w
            if best is not None and best_d is not None and best_d <= max_dist:
                return best.text
        return None

    # ------------------------------------------------------------------
    # Per-field extractors
    # ------------------------------------------------------------------
    def _extract_be_no_date(self, pages_rows) -> tuple[RawValue, RawValue]:
        """BE No (Req 2.1) and BE Date (Req 2.2) from the top ``Port Code | BE No |
        BE Date | BE Type`` band (repeated on every page)."""

        def is_be_band(row) -> bool:
            texts = [w.text.lower() for w in row]
            return (
                "port" in texts
                and "code" in texts
                and texts.count("be") >= 2
                and "no" in texts
                and "date" in texts
                and "type" in texts
            )

        found = self._find_label_row(pages_rows, is_be_band)
        if found is None:
            return RawValue.missing(), RawValue.missing()
        pi, ri, row = found
        rows = pages_rows[pi]

        i_no = self._word_where(row, lambda t: t.lower() == "no")
        i_date = self._word_where(row, lambda t: t.lower() == "date")

        be_no = RawValue.missing()
        be_date = RawValue.missing()
        if i_no is not None and i_no > 0:
            anchor = (row[i_no - 1].x0 + row[i_no].x1) / 2.0  # "BE" + "No" column
            be_no = self._capture(self._value_near_anchor(rows, ri, anchor))
        if i_date is not None and i_date > 0:
            anchor = (row[i_date - 1].x0 + row[i_date].x1) / 2.0  # "BE" + "Date"
            be_date = self._capture(self._value_near_anchor(rows, ri, anchor))
        return be_no, be_date

    def _extract_invoice_summary(
        self, pages_rows
    ) -> tuple[RawValue, RawValue, RawValue]:
        """Invoice No (Req 2.3), invoice amount and currency (Req 2.5) from the
        Part I summary band ``1.S.NO | 2.INVOICE NO | 3.INV. AMT | 4.CUR``."""

        def is_inv_summary(row) -> bool:
            texts = [w.text.lower() for w in row]
            return (
                any(t.startswith("2.invoice") for t in texts)
                and "3.inv." in texts
                and "4.cur" in texts
            )

        found = self._find_label_row(pages_rows, is_inv_summary)
        if found is None:
            return RawValue.missing(), RawValue.missing(), RawValue.missing()
        pi, ri, row = found
        rows = pages_rows[pi]

        inv_no = RawValue.missing()
        inv_amt = RawValue.missing()
        currency = RawValue.missing()

        i_inv = self._word_where(row, lambda t: t.lower().startswith("2.invoice"))
        if i_inv is not None and i_inv + 1 < len(row):
            anchor = (row[i_inv].x0 + row[i_inv + 1].x1) / 2.0  # "2.INVOICE" + "NO"
            inv_no = self._capture(self._value_near_anchor(rows, ri, anchor))

        i_amt = self._word_where(row, lambda t: t.lower() == "3.inv.")
        if i_amt is not None and i_amt + 1 < len(row):
            anchor = (row[i_amt].x0 + row[i_amt + 1].x1) / 2.0  # "3.INV." + "AMT"
            inv_amt = self._capture(
                self._value_near_anchor(rows, ri, anchor), numeric=True
            )

        i_cur = self._word_where(row, lambda t: t.lower() == "4.cur")
        if i_cur is not None:
            anchor = self._center(row[i_cur])
            currency = self._capture(self._value_near_anchor(rows, ri, anchor))

        return inv_no, inv_amt, currency

    def _extract_invoice_date(self, pages_rows) -> RawValue:
        """Invoice Date (Req 2.4) from the Part II ``2.INVOICE NO. & DT.`` column,
        where the date is printed on its own row below the invoice number."""

        def is_part2_inv(row) -> bool:
            texts = [w.text.lower() for w in row]
            return (
                any(t.startswith("2.invoice") for t in texts)
                and "dt." in texts
                and any(t.startswith("3.purchase") for t in texts)
            )

        found = self._find_label_row(pages_rows, is_part2_inv)
        if found is None:
            return RawValue.missing()
        pi, ri, row = found
        rows = pages_rows[pi]

        i_inv = self._word_where(row, lambda t: t.lower().startswith("2.invoice"))
        i_dt = self._word_where(row, lambda t: t.lower() == "dt.")
        if i_inv is None or i_dt is None:
            return RawValue.missing()
        x_lo = row[i_inv].x0 - 10.0
        x_hi = row[i_dt].x1 + 15.0

        # Scan the rows beneath the label for the first date-shaped token sitting
        # inside the invoice column (skips the invoice-number row above it).
        for rj in range(ri + 1, min(ri + 6, len(rows))):
            for w in rows[rj]:
                c = self._center(w)
                if x_lo <= c <= x_hi and _DATE_RE.fullmatch(w.text):
                    return self._capture(w.text)
        return RawValue.missing()

    def _extract_package_count(self, pages_rows) -> RawValue:
        """Total package/CTN count (Req 2.6) - the value printed immediately to
        the right of the ``PKG`` label in the ``BILL OF ENTRY ... PKG`` band."""
        found = self._find_label_row(
            pages_rows, lambda row: any(w.text.upper() == "PKG" for w in row)
        )
        if found is None:
            return RawValue.missing()
        _pi, _ri, row = found
        i_pkg = self._word_where(row, lambda t: t.upper() == "PKG")
        if i_pkg is None or i_pkg + 1 >= len(row):
            return RawValue.missing()
        return self._capture(row[i_pkg + 1].text, numeric=True)

    def _extract_party_name(self, pages_rows) -> RawValue:
        """Supplier/exporter Party Name (Req 2.7): the first line beneath the
        ``3.SUPPLIER NAME & ADDRESS`` label, taken from the left (supplier)
        column only so the adjacent Third Party column is never mixed in."""

        def is_supplier(row) -> bool:
            return any(w.text.lower().startswith("3.supplier") for w in row)

        found = self._find_label_row(pages_rows, is_supplier)
        if found is None:
            return RawValue.missing()
        pi, ri, row = found
        rows = pages_rows[pi]

        # Right edge of the supplier column = left edge of the Third Party column.
        i_third = self._word_where(row, lambda t: t.lower().startswith("4.third"))
        x_split = row[i_third].x0 - 10.0 if i_third is not None else 360.0

        if ri + 1 >= len(rows):
            return RawValue.missing()
        name_words = [w for w in rows[ri + 1] if w.x0 < x_split]
        if not name_words:
            return RawValue.missing()
        name_words.sort(key=lambda w: w.x0)
        return self._capture(" ".join(w.text for w in name_words))

    def _extract_container_details(self, pages_rows) -> RawValue:
        """Container details (Req 2.8): the container number(s) read under the
        ``5.CONTAINER NUMBER`` column, combined with the container count from the
        ``TYPE | INV | ITEM | CONT`` band."""
        numbers = self._extract_container_numbers(pages_rows)
        count = self._extract_container_count(pages_rows)

        if not numbers:
            return RawValue.missing()
        joined = ", ".join(numbers)
        if count is not None:
            return self._capture(f"{joined} (count: {count})")
        return self._capture(joined)

    def _extract_container_numbers(self, pages_rows) -> list[str]:
        """Collect the container number(s) printed beneath ``5.CONTAINER NUMBER``."""

        def is_container_hdr(row) -> bool:
            texts = [w.text.lower() for w in row]
            return any(t.startswith("5.container") for t in texts) and "number" in texts

        found = self._find_label_row(pages_rows, is_container_hdr)
        if found is None:
            return []
        pi, ri, row = found
        rows = pages_rows[pi]

        i_con = self._word_where(row, lambda t: t.lower().startswith("5.container"))
        i_num = self._word_where(row, lambda t: t.lower() == "number")
        if i_con is None or i_num is None:
            return []
        x_lo = row[i_con].x0 - 6.0
        x_hi = row[i_num].x1 + 8.0

        numbers: list[str] = []
        for rj in range(ri + 1, min(ri + 6, len(rows))):
            row_hit: str | None = None
            for w in rows[rj]:
                c = self._center(w)
                if x_lo <= c <= x_hi and re.fullmatch(r"[A-Za-z0-9]{6,}", w.text):
                    row_hit = w.text
                    break
            if row_hit is not None:
                numbers.append(row_hit)
            elif numbers:
                # Container rows are contiguous; stop once the block ends.
                break
        return numbers

    def _extract_container_count(self, pages_rows) -> str | None:
        """Container count from the ``TYPE | INV | ITEM | CONT`` band (the value
        aligned under the ``CONT`` label)."""

        def is_type_band(row) -> bool:
            texts = [w.text.upper() for w in row]
            return "TYPE" in texts and "INV" in texts and "ITEM" in texts and "CONT" in texts

        found = self._find_label_row(pages_rows, is_type_band)
        if found is None:
            return None
        pi, ri, row = found
        rows = pages_rows[pi]
        i_cont = self._word_where(row, lambda t: t.upper() == "CONT")
        if i_cont is None:
            return None
        anchor = self._center(row[i_cont])
        return self._value_near_anchor(rows, ri, anchor)

    def _extract_bl_no_date(self, pages_rows) -> tuple[RawValue, RawValue]:
        """Bill of Lading No/Date (Req 2.9), where present. Prefers the House B/L
        (``8.HAWB NO`` / ``9.DATE``) and falls back to the Master B/L
        (``6.MAWB NO`` / ``7.DATE``). Returns missing values (never flagged) when
        no B/L is present in the BOE."""
        for no_prefix, date_label in (("8.hawb", "9.date"), ("6.mawb", "7.date")):
            bl_no, bl_date = self._bl_from(pages_rows, no_prefix, date_label)
            if not bl_no.is_missing:
                return bl_no, bl_date
        return RawValue.missing(), RawValue.missing()

    def _bl_from(
        self, pages_rows, no_prefix: str, date_label: str
    ) -> tuple[RawValue, RawValue]:
        """Read one B/L number/date pair from the IGM band keyed by the given
        ``X.MAWB``/``X.HAWB`` number-label prefix and ``X.DATE`` date label."""

        def is_igm_band(row) -> bool:
            texts = [w.text.lower() for w in row]
            return any(t.startswith(no_prefix) for t in texts) and date_label in texts

        found = self._find_label_row(pages_rows, is_igm_band)
        if found is None:
            return RawValue.missing(), RawValue.missing()
        pi, ri, row = found
        rows = pages_rows[pi]

        bl_no = RawValue.missing()
        bl_date = RawValue.missing()

        # The IGM band carries its values on the single row directly below the
        # labels, so restrict the read to that row (max_scan=1); scanning further
        # would let an unrelated label row below supply a false value.
        i_no = self._word_where(row, lambda t: t.lower().startswith(no_prefix))
        if i_no is not None and i_no + 1 < len(row):
            anchor = (row[i_no].x0 + row[i_no + 1].x1) / 2.0  # "X.MAWB/HAWB" + "NO"
            text = self._value_near_anchor(rows, ri, anchor, max_scan=1)
            if text is not None:
                bl_no = self._capture(text)

        i_dt = self._word_where(row, lambda t: t.lower() == date_label)
        if i_dt is not None and not bl_no.is_missing:
            anchor = self._center(row[i_dt])
            text = self._value_near_anchor(rows, ri, anchor, max_scan=1)
            if text is not None:
                bl_date = self._capture(text)

        return bl_no, bl_date

    # ------------------------------------------------------------------
    # Line-item extraction (task 5.1)
    # ------------------------------------------------------------------
    def _capture_rate(self, raw_text: str | None) -> RawValue:
        """Capture a duty rate, interpreting the printed percentage as a fraction.

        The BOE prints duty rates as bare percentages (e.g. ``"18"`` for 18%,
        ``"15"`` for 15%) with no ``%`` sign. The verbatim characters are
        preserved in ``raw_text`` (Req 3.10/3.11) while ``parsed`` holds the
        decimal-fraction interpretation (``0.18``, ``0.15``) that the
        Value_Calculator and the target Excel use (Req 6.5; sample stores
        ``0.18``/``0.15``). An exemption-driven ``"0"`` becomes numeric ``0.0``
        (Req 3.14); blank -> missing; non-numeric -> unparseable (raw retained).
        """
        if raw_text is None or raw_text.strip() == "":
            return RawValue.missing()
        number = self._parse_number(raw_text)
        if number is None:
            return RawValue.unparseable(raw_text)
        return RawValue(
            raw_text=raw_text,
            parsed=number / 100.0,
            is_missing=False,
            is_unparseable=False,
        )

    def _token_in_range(
        self, row: list[Word], lo: float, hi: float
    ) -> str | None:
        """Text of the token in ``row`` whose centre lies in ``[lo, hi]``.

        When several tokens fall inside the window the one closest to its centre
        is returned; ``None`` when the column is empty so a blank cell never
        grabs an unrelated neighbour.
        """
        mid = (lo + hi) / 2.0
        best: Word | None = None
        best_d: float | None = None
        for w in row:
            c = self._center(w)
            if lo <= c <= hi:
                d = abs(c - mid)
                if best_d is None or d < best_d:
                    best_d, best = d, w
        return best.text if best is not None else None

    # -- Part II: invoice items ----------------------------------------
    def _extract_invoice_items(self, pages) -> dict[int, InvoiceItemRow]:
        """Extract the Part II invoice/valuation table for every line item.

        Returns ``{item_serial: InvoiceItemRow}`` carrying the CTH/HSN code,
        the complete (wrapped) description, unit price, quantity, UQC and the
        line amount, each captured verbatim (Req 3.4-3.8). Reads the
        orientation-aware upright-word rows so the BOE's rotated margin labels
        never contaminate the table; column membership is decided positionally
        from the table header on each page.
        """
        working: dict[int, dict] = {}
        for page in pages:
            rows = self._reconstruct_rows(self._upright_words(page))
            self._invoice_items_from_page(rows, working)
        return {serial: self._build_invoice_row(entry) for serial, entry in working.items()}

    def _find_invoice_header(self, rows) -> tuple[int, _InvCols] | None:
        """Locate the Part II table header row and derive its column splits."""
        for ri, row in enumerate(rows):
            low = [w.text.lower() for w in row]
            if (
                "2.cth" in low
                and "3.description" in low
                and "5.quantity" in low
                and "7.amount" in low
            ):
                cols = self._invoice_splits(row)
                if cols is not None:
                    return ri, cols
        return None

    @staticmethod
    def _x0_of(row, predicate) -> float | None:
        for w in row:
            if predicate(w.text):
                return w.x0
        return None

    @staticmethod
    def _x1_of(row, predicate) -> float | None:
        for w in row:
            if predicate(w.text):
                return w.x1
        return None

    def _invoice_splits(self, hrow) -> _InvCols | None:
        """Compute Part II column boundaries from the header label positions."""
        unit = self._x0_of(hrow, lambda t: t.lower() == "4.unit")
        qty = self._x0_of(hrow, lambda t: t.lower() == "5.quantity")
        uqc0 = self._x0_of(hrow, lambda t: t.lower() == "6.uqc")
        uqc1 = self._x1_of(hrow, lambda t: t.lower() == "6.uqc")
        amt = self._x0_of(hrow, lambda t: t.lower() == "7.amount")
        if None in (unit, qty, uqc0, uqc1, amt):
            return None
        return _InvCols(
            desc_up=unit,
            up_qty=qty,
            qty_uqc=uqc0,
            uqc_amt=(uqc1 + amt) / 2.0,
        )

    def _invoice_items_from_page(self, rows, working: dict) -> None:
        """Parse one page's Part II rows into ``working`` (keyed by serial)."""
        header = self._find_invoice_header(rows)
        if header is None:
            return
        ri0, cols = header
        current: dict | None = None
        for ri in range(ri0 + 1, len(rows)):
            row = rows[ri]
            if any(w.text.upper() == "GLOSSARY" for w in row):
                break  # end of the invoice table on this page
            if self._is_invoice_footer(row):
                break
            if self._is_invoice_item_row(row, cols):
                entry = self._parse_invoice_item_row(row, cols)
                if entry is None:
                    continue
                working[entry["serial"]] = entry
                current = entry
            elif current is not None:
                # Description continuation line: append its left-column tokens.
                current["desc"].extend(self._left_tokens(row, cols))

    @staticmethod
    def _is_invoice_footer(row) -> bool:
        """True for page-footer rows (``Page x Of 30`` / ``Verify using ...``)."""
        texts = [w.text for w in row]
        if texts[:1] == ["Verify"]:
            return True
        low = [t.lower() for t in texts]
        return "page" in low and "of" in low and len(texts) <= 5

    def _left_tokens(self, row, cols: _InvCols) -> list[str]:
        """Description-region tokens of ``row`` (centre left of the unit-price column)."""
        left = [w for w in row if self._center(w) < cols.desc_up]
        left.sort(key=lambda w: w.x0)
        return [w.text for w in left]

    def _is_invoice_item_row(self, row, cols: _InvCols) -> bool:
        """A new item row: leftmost token is an integer serial and the right-hand
        value columns (unit price / qty / amount) carry tokens."""
        left = [w for w in row if self._center(w) < cols.desc_up]
        if not left:
            return False
        left.sort(key=lambda w: w.x0)
        if not re.fullmatch(r"\d+", left[0].text):
            return False
        return any(self._center(w) >= cols.desc_up for w in row)

    def _parse_invoice_item_row(self, row, cols: _InvCols) -> dict | None:
        """Capture one item's invoice fields; description starts as this row's
        left tokens (continuation lines are appended later)."""
        left = sorted(
            (w for w in row if self._center(w) < cols.desc_up), key=lambda w: w.x0
        )
        if not left or not re.fullmatch(r"\d+", left[0].text):
            return None
        serial = int(left[0].text)
        cth = left[1].text if len(left) > 1 else None
        desc_tokens = [w.text for w in left[2:]]

        unit_price = self._token_in_range(row, cols.desc_up, cols.up_qty)
        qty = self._token_in_range(row, cols.up_qty, cols.qty_uqc)
        uqc = self._token_in_range(row, cols.qty_uqc, cols.uqc_amt)
        amount = self._token_in_range(row, cols.uqc_amt, float("inf"))

        return {
            "serial": serial,
            "cth": cth,
            "desc": desc_tokens,
            "unit_price": unit_price,
            "qty": qty,
            "uqc": uqc,
            "amount": amount,
        }

    def _build_invoice_row(self, entry: dict) -> InvoiceItemRow:
        """Convert a working entry into an immutable :class:`InvoiceItemRow`."""
        description = " ".join(t for t in entry["desc"] if t).strip()
        return InvoiceItemRow(
            item_serial=entry["serial"],
            cth_hsn=self._capture(entry["cth"]),
            description=self._capture(description if description else None),
            unit_price_usd=self._capture(entry["unit_price"], numeric=True),
            quantity=self._capture(entry["qty"], numeric=True),
            unit=self._capture(entry["uqc"]),
            amount=self._capture(entry["amount"], numeric=True),
        )

    # -- Part III: duty items ------------------------------------------
    def _extract_duty_items(self, pages) -> dict[int, DutyItemRow]:
        """Extract the Part III per-item duty blocks for every line item.

        Returns ``{item_serial: DutyItemRow}`` carrying the assessable value,
        BCD rate/amount, SWS amount, IGST rate and total duty (Req 3.9-3.12).
        Each item is one positional block keyed by ``2.ITEMSN``; the duty grid's
        ``Rate``/``Amount`` sub-rows are read by column window. Exemption-driven
        zero rates/amounts are captured as numeric ``0`` (Req 3.14).
        """
        duties: dict[int, DutyItemRow] = {}
        for page in pages:
            rows = self._reconstruct_rows(self._upright_words(page))
            self._duty_items_from_page(rows, duties)
        return duties

    @staticmethod
    def _is_invsno_header(row) -> bool:
        low = [w.text.lower() for w in row]
        return "1.invsno" in low and "2.itemsn" in low

    def _duty_items_from_page(self, rows, duties: dict) -> None:
        """Split a page into per-item duty blocks (one per ``1.INVSNO`` header)."""
        starts = [ri for ri, row in enumerate(rows) if self._is_invsno_header(row)]
        for k, start in enumerate(starts):
            end = starts[k + 1] if k + 1 < len(starts) else len(rows)
            block = rows[start:end]
            result = self._parse_duty_block(block)
            if result is not None:
                serial, duty = result
                duties[serial] = duty

    def _parse_duty_block(self, block) -> tuple[int, DutyItemRow] | None:
        """Parse a single Part III item block into a :class:`DutyItemRow`."""
        if len(block) < 2:
            return None
        header = block[0]
        itemsn_x0 = self._x0_of(header, lambda t: t.lower() == "2.itemsn")
        cth_x0 = self._x0_of(header, lambda t: t.lower() == "3.cth")
        if itemsn_x0 is None or cth_x0 is None:
            return None
        serial_text = self._token_in_range(block[1], itemsn_x0 - 2, cth_x0 - 3)
        if serial_text is None or not re.fullmatch(r"\d+", serial_text):
            return None
        serial = int(serial_text)

        assessable, total_duty = self._duty_assess_total(block)
        bcd_rate, bcd_amount, sws_amount, igst_rate = self._duty_grid(block)

        return serial, DutyItemRow(
            item_serial=serial,
            assessable_value=assessable,
            bcd_rate=bcd_rate,
            bcd_amount=bcd_amount,
            sws_amount=sws_amount,
            igst_rate=igst_rate,
            total_duty=total_duty,
        )

    def _duty_assess_total(self, block) -> tuple[RawValue, RawValue]:
        """Assessable value (29.ASSESS VALUE) and total duty (30. TOTAL DUTY)."""
        for bi in range(len(block) - 1):
            row = block[bi]
            low = [w.text.lower() for w in row]
            if "29.assess" in low and "30." in low:
                assess_x0 = self._x0_of(row, lambda t: t.lower() == "29.assess")
                total_x0 = self._x0_of(row, lambda t: t.lower() == "30.")
                data = block[bi + 1]
                assess = self._capture(
                    self._token_in_range(data, assess_x0 - 10, total_x0 - 6),
                    numeric=True,
                )
                total = self._capture(
                    self._token_in_range(data, total_x0 - 6, float("inf")),
                    numeric=True,
                )
                return assess, total
        return RawValue.missing(), RawValue.missing()

    def _duty_grid(self, block) -> tuple[RawValue, RawValue, RawValue, RawValue]:
        """BCD rate/amount, SWS amount and IGST rate from the first duty grid.

        The grid carries one ``Rate`` and one ``Amount`` sub-row beneath a header
        of duty-type columns (``BCD``, ``3.SWS``, ``5.IGST``, ...). Values are
        read by the column window each label defines. Returns missing values if
        the grid or its sub-rows cannot be located.
        """
        di = self._find_duty_grid_header(block)
        if di is None:
            return (RawValue.missing(),) * 4  # type: ignore[return-value]
        cols = self._duty_grid_cols(block[di])
        if cols is None:
            return (RawValue.missing(),) * 4  # type: ignore[return-value]

        rate_row = self._first_row_starting(block, di + 1, "rate")
        amount_row = self._first_row_starting(block, di + 1, "amount")

        bcd_rate = (
            self._capture_rate(self._token_in_range(rate_row, *cols.bcd))
            if rate_row is not None
            else RawValue.missing()
        )
        igst_rate = (
            self._capture_rate(self._token_in_range(rate_row, *cols.igst))
            if rate_row is not None
            else RawValue.missing()
        )
        bcd_amount = (
            self._capture(self._token_in_range(amount_row, *cols.bcd), numeric=True)
            if amount_row is not None
            else RawValue.missing()
        )
        sws_amount = (
            self._capture(self._token_in_range(amount_row, *cols.sws), numeric=True)
            if amount_row is not None
            else RawValue.missing()
        )
        return bcd_rate, bcd_amount, sws_amount, igst_rate

    @staticmethod
    def _find_duty_grid_header(block) -> int | None:
        """Index of the first duty grid header (the one carrying BCD/SWS/IGST)."""
        for bi, row in enumerate(block):
            low = [w.text.lower() for w in row]
            if "bcd" in low and "3.sws" in low and "5.igst" in low:
                return bi
        return None

    def _duty_grid_cols(self, header) -> _DutyCols | None:
        """Column windows for BCD/SWS/IGST values from the duty grid header."""
        cvd = self._x0_of(header, lambda t: t.lower().startswith("2.cvd"))
        sws = self._x0_of(header, lambda t: t.lower() == "3.sws")
        sad = self._x0_of(header, lambda t: t.lower() == "4.sad")
        igst = self._x0_of(header, lambda t: t.lower() == "5.igst")
        gcess = self._x0_of(header, lambda t: t.lower().startswith("6.g"))
        if None in (cvd, sws, sad, igst, gcess):
            return None
        return _DutyCols(
            bcd=(100.0, cvd - 3.0),
            sws=(sws - 14.0, sad - 3.0),
            igst=(igst - 14.0, gcess - 3.0),
        )

    @staticmethod
    def _first_row_starting(block, start: int, first_word: str):
        """First row at/after ``start`` whose leftmost token equals ``first_word``."""
        for bi in range(start, len(block)):
            row = block[bi]
            if row and row[0].text.lower() == first_word:
                return row
        return None

    # ------------------------------------------------------------------
    # Declared item count (Req 3.2)
    # ------------------------------------------------------------------
    def _extract_declared_item_count(self, pages) -> int | None:
        """The BOE's declared total item count (Req 3.2).

        Read from the ``TYPE | INV | ITEM | CONT`` band: ``ITEM`` carries the
        document's declared number of line items (45 for the reference BOE). The
        value aligned under the ``ITEM`` label is taken positionally. Returns
        ``None`` when the band or a clean integer value cannot be located, so the
        orchestrator can surface "could not verify" rather than a wrong count.
        """
        pages_rows = [
            self._reconstruct_rows(self._upright_words(page)) for page in pages
        ]

        def is_type_band(row) -> bool:
            texts = [w.text.upper() for w in row]
            return (
                "TYPE" in texts
                and "INV" in texts
                and "ITEM" in texts
                and "CONT" in texts
            )

        found = self._find_label_row(pages_rows, is_type_band)
        if found is None:
            return None
        pi, ri, row = found
        rows = pages_rows[pi]
        i_item = self._word_where(row, lambda t: t.upper() == "ITEM")
        if i_item is None:
            return None
        anchor = self._center(row[i_item])
        text = self._value_near_anchor(rows, ri, anchor)
        if text is None or not re.fullmatch(r"\d+", text.strip()):
            return None
        return int(text.strip())

    # ------------------------------------------------------------------
    # Two-pass assembly keyed by item serial (task 5.2)
    # ------------------------------------------------------------------
    def _merge_items(
        self,
        inv: dict[int, "InvoiceItemRow"],
        duty: dict[int, "DutyItemRow"],
        declared_count: int | None,
    ) -> tuple[list[LineItem], list[ReviewFlag]]:
        """Join Part II (invoice) and Part III (duty) rows on ``item_serial``.

        Builds one :class:`LineItem` per serial present in *either* source, in
        ascending serial order. Because ``_extract_invoice_items`` /
        ``_extract_duty_items`` already collapse each item's (possibly
        multi-page) block into a single row keyed by serial, every field value
        appears exactly once here -- multi-page items are stitched without
        duplication or omission (Req 3.13).

        No serial is ever dropped: a serial found only in Part II (or only in
        Part III) still yields a complete ``LineItem`` whose fields from the
        absent source are :meth:`RawValue.missing` (Req 3.1, 9.5). Every required
        field (Req 3.4-3.12) that is missing or unparseable produces a
        ``ReviewFlag(line_item, serial, field, MISSING/UNPARSEABLE)`` retaining
        the raw text for the unparseable case (Req 3.15, 9.2, 9.3); no inferred
        or default value is ever substituted.

        ``declared_count`` is accepted for interface completeness (the design's
        signature) and threaded through by ``parse`` onto the
        ``ExtractedDocument``; the declared-vs-extracted count cross-check itself
        is the orchestrator's responsibility (Req 3.3, task 9.3).
        """
        serials = sorted(set(inv) | set(duty))
        line_items: list[LineItem] = []
        flags: list[ReviewFlag] = []

        for serial in serials:
            inv_row = inv.get(serial)
            duty_row = duty.get(serial)

            cth_hsn = inv_row.cth_hsn if inv_row else RawValue.missing()
            description = inv_row.description if inv_row else RawValue.missing()
            unit_price_usd = inv_row.unit_price_usd if inv_row else RawValue.missing()
            quantity = inv_row.quantity if inv_row else RawValue.missing()
            unit = inv_row.unit if inv_row else RawValue.missing()

            assessable_value = (
                duty_row.assessable_value if duty_row else RawValue.missing()
            )
            bcd_rate = duty_row.bcd_rate if duty_row else RawValue.missing()
            bcd_amount = duty_row.bcd_amount if duty_row else RawValue.missing()
            igst_rate = duty_row.igst_rate if duty_row else RawValue.missing()
            total_duty = duty_row.total_duty if duty_row else RawValue.missing()

            line_items.append(
                LineItem(
                    item_serial=serial,
                    cth_hsn=cth_hsn,
                    description=description,
                    unit_price_usd=unit_price_usd,
                    quantity=quantity,
                    unit=unit,
                    assessable_value=assessable_value,
                    bcd_rate=bcd_rate,
                    bcd_amount=bcd_amount,
                    igst_rate=igst_rate,
                    total_duty=total_duty,
                )
            )

            # Required fields per Req 3.4-3.12 (item_serial is the always-present
            # int join key and so is not flaggable here).
            required: list[tuple[str, RawValue]] = [
                ("cth_hsn", cth_hsn),
                ("description", description),
                ("unit_price_usd", unit_price_usd),
                ("quantity", quantity),
                ("unit", unit),
                ("assessable_value", assessable_value),
                ("bcd_rate", bcd_rate),
                ("bcd_amount", bcd_amount),
                ("igst_rate", igst_rate),
                ("total_duty", total_duty),
            ]
            for field_name, value in required:
                if value.is_missing:
                    flags.append(
                        ReviewFlag(
                            scope="line_item",
                            field_name=field_name,
                            reason="MISSING",
                            item_serial=serial,
                        )
                    )
                elif value.is_unparseable:
                    flags.append(
                        ReviewFlag(
                            scope="line_item",
                            field_name=field_name,
                            reason="UNPARSEABLE",
                            item_serial=serial,
                            raw_text=value.raw_text,
                        )
                    )

        return line_items, flags
