"""Invoice / Packing-List parser: per-line carton (CTN) counts.

The Bill of Entry PDF does not carry a per-line carton count (only a document
level total). The supplier's *commercial invoice* does: its line table has a
``TOTAL CTNS`` column keyed by ``SR NO`` that maps 1:1 to the BOE line serial.

This parser reads that ``TOTAL CTNS`` column from the **invoice** pages of an
"invoice cum packing list" PDF and returns a ``{serial: cartons}`` mapping that
the orchestrator attaches to the matching BOE line items (so the Excel CTN
column, ``G``, is populated per line). It is intentionally tolerant: cells that
are blank in the invoice yield no entry (that line stays blank), and the more
granular *packing list* pages (a different SL-NO breakdown that would not align
1:1 with the BOE) are skipped.

Extraction is positional: the ``TOTAL CTNS`` column sits at a stable horizontal
band, so a blank carton cell is simply the absence of a token in that band -
which a text-only parse could not distinguish from the quantity column.
"""

from __future__ import annotations

import re

from boe_converter.models import RawValue

# Units seen in the invoice's quantity column; used to recognize a data row and
# to bound the carton column on its right.
_UNIT_RE = re.compile(r"^[A-Za-z]{2,4}$")
_UNITS = {
    "PCS", "PC", "DOZ", "KGS", "KG", "GRS", "THD", "SET", "NOS", "EA", "UNT",
    "MTR", "PRS", "BOX", "MTS", "MT", "CTN", "PKT", "BAG", "ROL", "YDS",
    "LTR", "SQM", "PAIR",
}

# Horizontal half-width (points) of the carton column band around the ``CTNS``
# header token's centre. The invoice's CTNS values sit within ~15pt of the
# header centre; the quantity column is ~55pt to the right, well outside.
_CTNS_BAND = 22.0

# Header spellings accepted for the carton column (case-insensitive, trailing
# punctuation ignored): suppliers print TOTAL CTNS, TOTAL CTN, CTNS. etc.
_CTNS_TOKENS = {"CTN", "CTNS"}

# Words expected in the table's header row alongside the carton header. Used to
# prefer the real column header over a stray "CTNS" in title/total text.
_HEADER_HINTS = {"SR", "NO", "DESCRIPTION", "QTY", "PRICE", "AMOUNT", "UNIT", "UQC"}

# Max gap (points) between two in-band digit tokens for them to count as one
# value split by the extractor (e.g. "20" read as "2" + "0"). pdfplumber can
# split wide-tracked digits into separate words; joining only near-adjacent
# tokens keeps two genuinely separate numbers from merging.
_DIGIT_JOIN_GAP = 6.0

# Max vertical gap (points) for stitching a serial-bearing row fragment to an
# adjacent unit-bearing fragment of the same logical invoice line. Covers
# baseline drift that splits one line across two geometric rows; normal row
# pitch (~15pt+) stays well outside it.
_ROW_STITCH_GAP = 8.0

# Header words locating the invoice value columns, mapped to the detail key
# they fill. The header prints them inside longer labels (``TOTAL QTY``,
# ``UNIT PRICE``, ``TOTAL AMOUNT``), so single-word matching suffices.
_VALUE_TOKENS = {"QTY": "quantity", "PRICE": "unit_price", "AMOUNT": "amount"}

# Horizontal half-width (points) of each value column band around its header
# token's centre. Value columns sit a full column apart, so this never reaches
# a neighbouring column's numbers; the unit column between QTY and PRICE holds
# no numerics (fused ``4000THD`` splits are centred on it, outside every band).
_VALUE_BAND = 20.0

# Words marking a totals/footer row: never absorb such a row while stitching,
# so a footer ("TOTAL CTN : 1159") cannot donate its number to a serial above.
_NON_DATA_HINTS = {"TOTAL", "SUBTOTAL", "GRAND", "BALANCE"}


def _header_token(text: str) -> str | None:
    """The carton-header token in ``text``, else ``None``."""
    cleaned = text.strip().upper().strip(".,:;")
    return cleaned if cleaned in _CTNS_TOKENS else None

# A row's SR NO sits in the far-left column.
_SR_NO_MAX_X = 70.0


def _serial_text(text: str) -> int | None:
    """A serial number from a far-left token (tolerating ``1.``), else ``None``."""
    cleaned = text.strip().rstrip(".").strip()
    return int(cleaned) if cleaned.isdigit() else None


def _xcenter(w: dict) -> float:
    """Horizontal centre of an extractor word."""
    return (float(w["x0"]) + float(w["x1"])) / 2.0


def _unit_token(text: str) -> str | None:
    """A known unit carried by ``text``, else ``None``.

    Accepts the bare token (``THD``) plus punctuated/merged extractions the PDF
    reader produces (``THD.``, ``CTN:``, ``THD$0.12``): a leading 2-4 letter run
    in the unit table counts, since tight supplier layouts merge the unit with
    its right-hand neighbour. A token starting with a digit (``2PCS``,
    ``12CM``) is never matched here; a fused ``number+unit`` (``4000THD``) is
    :func:`_fused_unit`'s job.
    """
    cleaned = text.strip().strip(".,:;")
    if _UNIT_RE.match(cleaned) and cleaned.upper() in _UNITS:
        return cleaned.upper()
    m = re.match(r"^([A-Za-z]{2,4})[^A-Za-z\s]", cleaned)
    if m and m.group(1).upper() in _UNITS:
        return m.group(1).upper()
    return None


def _fused_unit(text: str) -> str | None:
    """The unit part of a fused ``number+unit`` token (``4000THD``), else ``None``."""
    core = text.strip().rstrip(".,:;")
    m = re.fullmatch(r"\d+(?:\.\d+)?([A-Za-z]{2,4})", core)
    if m and m.group(1).upper() in _UNITS:
        return m.group(1).upper()
    return None


def _split_merged_token(w: dict) -> list[dict]:
    """Split a fused ``number+unit`` token for matching (geometry shared).

    Only the returned pieces drive the serial/unit/carton checks; captured
    description text always uses the original word, so names like
    ``126PCS STATIONERY SET`` stay verbatim.
    """
    unit = _fused_unit(w["text"])
    if unit is None:
        return [w]
    core = w["text"].strip().rstrip(".,:;")
    num = core[: len(core) - len(unit)]
    width = float(w["x1"]) - float(w["x0"])
    cut = float(w["x0"]) + width * len(num) / (len(num) + len(unit))
    return [
        {**w, "text": num, "x1": cut},
        {**w, "text": unit, "x0": cut},
    ]


def _row_top(row: list[dict]) -> float:
    """Top edge of a row group (its highest word)."""
    return min(float(w["top"]) for w in row)


def _row_has_sr(row: list[dict]) -> bool:
    """True when the row carries a far-left serial number."""
    return any(
        _xcenter(w) < _SR_NO_MAX_X and _serial_text(w["text"]) is not None
        for w in row
    )


def _row_has_unit(row: list[dict]) -> bool:
    """True when the row carries any known unit token (fused ones count)."""
    return any(
        _unit_token(w["text"]) is not None or _fused_unit(w["text"]) is not None
        for w in row
    )


def _row_has_totals_word(row: list[dict]) -> bool:
    """True when the row looks like a totals/footer row."""
    return any(
        w["text"].strip().upper().strip(".,:;") in _NON_DATA_HINTS for w in row
    )


def _stitch_split_rows(rows: list[list[dict]]) -> list[list[dict]]:
    """Rejoin logical invoice lines split across two geometric rows.

    When a row carries a serial but no unit (baseline drift pushed its numeric
    half out of the group), it absorbs the adjacent unit-bearing, serial-free
    group within ``_ROW_STITCH_GAP``. Groups with both halves intact, and
    totals/footer rows, pass through untouched.
    """
    stitched: list[list[dict]] = []
    consumed: set[int] = set()
    for i, row in enumerate(rows):
        if i in consumed:
            continue
        if _row_has_sr(row) and not _row_has_unit(row):
            partner = None
            for j in (i + 1, i - 1):
                if not 0 <= j < len(rows) or j in consumed:
                    continue
                other = rows[j]
                if _row_has_sr(other) or not _row_has_unit(other):
                    continue
                if _row_has_totals_word(other):
                    continue
                if abs(_row_top(other) - _row_top(row)) > _ROW_STITCH_GAP:
                    continue
                partner = j
                break
            if partner is not None:
                if partner == i - 1:
                    row = stitched.pop() + row
                else:
                    row = row + rows[partner]
                consumed.add(partner)
        stitched.append(row)
    return stitched


class InvoicePackingListParser:
    """Extracts per-line carton counts from an invoice / packing-list PDF."""

    ROW_TOLERANCE = 3.0

    def parse_cartons(self, doc) -> dict[int, RawValue]:
        """Return ``{serial: cartons}`` read from the invoice pages of ``doc``.

        Backward-compatible wrapper over :meth:`parse_line_details`.
        """
        return {
            serial: detail["cartons"]
            for serial, detail in self.parse_line_details(doc).items()
            if detail.get("cartons") is not None
        }

    def parse_line_details(self, doc) -> dict[int, dict]:
        """Return ``{serial: {...}}`` per line.

        Reads the invoice line table (packing-list pages and the totals row are
        skipped). For each line the per-line carton count (``TOTAL CTNS`` column),
        the ``DESCRIPTION OF GOODS`` text, and the ``TOTAL QTY`` / ``UNIT PRICE``
        / ``TOTAL AMOUNT`` values are captured, keyed by ``SR NO`` (which maps
        1:1 to the BOE line serial). Blank cells yield ``None`` for that field.
        ``doc`` may be a path, bytes buffer, or an opened ``pdfplumber.PDF``.
        """
        handle, pages, should_close = self._resolve(doc)
        try:
            details: dict[int, dict] = {}
            for page in pages:
                text = page.extract_text() or ""
                if not self._is_invoice_page(text):
                    continue
                self._parse_page(page, details)
            return details
        finally:
            if should_close and handle is not None:
                try:
                    handle.close()
                except Exception:
                    pass

    # ------------------------------------------------------------------
    @staticmethod
    def _resolve(doc):
        if hasattr(doc, "pages"):
            return doc, list(doc.pages), False
        try:
            import pdfplumber
        except ImportError as exc:
            raise ValueError(
                "pdfplumber is required to read an invoice PDF; "
                "install it (pip install pdfplumber) or pass an opened "
                "document with a .pages attribute."
            ) from exc
        handle = pdfplumber.open(doc)
        return handle, list(handle.pages), True

    @staticmethod
    def _is_invoice_page(text: str) -> bool:
        """True for an invoice line-table page (not a packing-list page).

        The invoice header carries ``UNIT PRICE`` / ``TOTAL AMOUNT``; the packing
        list carries ``QTY PER CTNS`` / ``PACKINGLIST`` and a different SL-NO
        breakdown that must not be read as per-line cartons.
        """
        low = text.lower()
        if "packinglist" in low or "qty per" in low:
            return False
        return "amount" in low and "price" in low

    def _rows(self, page) -> list[list[dict]]:
        """Group the page's words into geometric rows (top-to-bottom)."""
        words = page.extract_words()
        words.sort(key=lambda w: (round(float(w["top"]), 1), float(w["x0"])))
        rows: list[list[dict]] = []
        current: list[dict] = []
        cur_top: float | None = None
        for w in words:
            top = float(w["top"])
            if cur_top is None or abs(top - cur_top) <= self.ROW_TOLERANCE:
                current.append(w)
                cur_top = top if cur_top is None else (cur_top + top) / 2.0
            else:
                rows.append(current)
                current = [w]
                cur_top = top
        if current:
            rows.append(current)
        return rows

    @staticmethod
    def _center(w: dict) -> float:
        return _xcenter(w)

    def _ctns_header(self, rows: list[list[dict]]) -> tuple[int, float] | None:
        """Locate the carton column header: ``(row index, horizontal centre)``.

        Accepts ``CTN``/``CTNS`` in any case with trailing punctuation (suppliers
        vary the label). When several candidates exist, the one sharing its row
        with table-header words (SR/NO/DESCRIPTION/QTY/...) wins, so a stray
        "TOTAL CTNS: 500" in title text cannot hijack the band; otherwise the
        first candidate is used.

        A header has line items beneath it, so only rows above the page's last
        line item are candidates - a footer ("TOTAL CTNS: 1159 TOTAL QTY 4500")
        cannot displace a header printed on a row of its own. A row carrying a
        serial is a line item, never the header: ``CTN`` is also a unit, so a
        data row would otherwise qualify.
        """
        last_item = max(
            (i for i, row in enumerate(rows) if _row_has_sr(row) and _row_has_unit(row)),
            default=0,
        )
        hinted: list[tuple[int, float]] = []
        plain: list[tuple[int, float]] = []
        for index, row in enumerate(rows[:last_item]):
            if _row_has_sr(row):
                continue
            hits = [w for w in row if _header_token(w["text"])]
            if not hits:
                continue
            row_words = {
                w["text"].strip().upper().strip(".,:;") for w in row
            }
            found = (index, self._center(hits[0]))
            (hinted if row_words & _HEADER_HINTS else plain).append(found)
        candidates = hinted + plain
        return candidates[0] if candidates else None

    def _value_columns(self, header_row: list[dict]) -> dict[str, float]:
        """Map detail key -> column centre for the invoice value columns.

        Looks for the ``QTY`` / ``PRICE`` / ``AMOUNT`` header words in the same
        header row that carries the carton header (suppliers print ``TOTAL
        QTY``, ``UNIT PRICE``, ``TOTAL AMOUNT``). A column whose header word is
        absent simply yields no values - the caller then leaves that field
        blank rather than guessing from a neighbouring column.
        """
        centers: dict[str, float] = {}
        for w in header_row:
            key = _VALUE_TOKENS.get(w["text"].strip().upper().strip(".,:;"))
            if key is not None and key not in centers:
                centers[key] = self._center(w)
        return centers

    def _column_value(self, gate: list[dict], center: float) -> RawValue | None:
        """The money value under a value-column centre, else ``None``.

        The nearest in-band numeric run wins (digit-split tokens are joined
        first); amounts print with currency symbols and thousand separators
        (``$1,500.00``), which are stripped before parsing. Anything
        non-numeric yields ``None`` - a missing value is never fabricated.
        """
        nearby = [
            w for w in gate
            if center - _VALUE_BAND <= self._center(w) <= center + _VALUE_BAND
            and self._is_money(w["text"])
        ]
        if not nearby:
            return None
        return self._money_value(self._join_number_tokens(nearby))

    def _parse_page(self, page, details: dict[int, dict]) -> None:
        rows = self._rows(page)
        rows = _stitch_split_rows(rows)
        header = self._ctns_header(rows)
        if header is None:
            return
        header_index, ctns_center = header
        value_centers = self._value_columns(rows[header_index])
        lo, hi = ctns_center - _CTNS_BAND, ctns_center + _CTNS_BAND
        # The description column sits between the SR NO column and the CTNS band.
        desc_hi = ctns_center - _CTNS_BAND

        # The table sits below its column header; anything above it (letterhead,
        # addresses) is not a line item even when it starts with a number.
        for row in rows[header_index + 1:]:
            ordered = sorted(row, key=lambda w: float(w["x0"]))
            # Matching sees through fused number+unit tokens (4000THD); the
            # captured description below always uses the original words.
            gate = [s for w in ordered for s in _split_merged_token(w)]
            # SR NO: an integer in the far-left column (tolerating a trailing
            # dot, e.g. "1." as some invoices print it).
            sr_word = next(
                (w for w in gate
                 if self._center(w) < _SR_NO_MAX_X and _serial_text(w["text"]) is not None),
                None,
            )
            if sr_word is None:
                continue
            # Require a unit token (e.g. PCS/DOZ) so totals/footer rows are
            # ignored - they have a serial-like number but no unit. The match
            # is punctuation-tolerant (``THD.``, ``CTN:``, ``THD$0.12``) since
            # tight supplier layouts merge the unit with its neighbour.
            if not _row_has_unit(gate):
                continue

            serial = _serial_text(sr_word["text"])
            assert serial is not None
            # A serial is captured once. A numbered note below the table
            # ("1. PACKING: 20 PCS PER CTN") reads as serial 1 with a unit; it
            # must not replace the real line 1, on this page or an earlier one.
            if serial in details:
                continue
            entry = details[serial] = {"cartons": None, "description": None}

            # Carton count: numeric token(s) whose centre falls in the CTNS
            # band. Adjacent digit tokens are joined first: extractors can
            # split "20" into "2" + "0", and taking the first token turned
            # 20 into 2. Only near-adjacent tokens merge, so genuinely
            # separate numbers never combine.
            ctn_words = [
                w for w in gate
                if lo <= self._center(w) <= hi
                and self._is_number(w["text"])
            ]
            if ctn_words:
                entry["cartons"] = self._carton_value(
                    self._join_number_tokens(ctn_words)
                )

            # TOTAL QTY / UNIT PRICE / TOTAL AMOUNT under their own header
            # columns (located once per page above). A column whose header word
            # was absent, or a blank cell, leaves that field None - a missing
            # value is never fabricated, and without both sides numeric the
            # orchestrator simply skips that field's comparison.
            for field, center in value_centers.items():
                entry[field] = self._column_value(gate, center)

            # Description: all tokens between the SR NO column and the CTNS band,
            # left-to-right (e.g. "SLIDERS (GARMENT ACCESSORY)"). Original
            # words are used so fused names ("126PCS STATIONERY SET") stay
            # verbatim.
            desc_words = [
                w for w in ordered
                if _SR_NO_MAX_X <= self._center(w) < desc_hi
            ]
            desc = " ".join(w["text"].strip() for w in desc_words).strip()
            if desc:
                entry["description"] = RawValue(raw_text=desc, parsed=desc)

    @staticmethod
    def _is_number(text: str) -> bool:
        return bool(re.fullmatch(r"\d+(?:\.\d+)?", text.strip().replace(",", "")))

    @staticmethod
    def _is_money(text: str) -> bool:
        """True for a money/quantity token (``4000``, ``0.12``, ``$1,500.00``)."""
        cleaned = text.strip().replace(",", "").lstrip("$₹€£ ").strip()
        return bool(re.fullmatch(r"\d+(?:\.\d+)?", cleaned))

    @staticmethod
    def _money_value(text: str) -> RawValue:
        """Wrap a money token verbatim, parsing it as an int when whole."""
        cleaned = text.strip().replace(",", "").lstrip("$₹€£ ").strip()
        try:
            number = float(cleaned)
        except ValueError:
            return RawValue(raw_text=text.strip(), parsed=text.strip())
        parsed: float | int = int(number) if number.is_integer() else number
        return RawValue(raw_text=text.strip(), parsed=parsed)

    @staticmethod
    def _join_number_tokens(words: list[dict]) -> str:
        """Join in-band number tokens split by the extractor into one value.

        Tokens are ordered left-to-right and merged while the gap between one
        token's end and the next token's start stays within
        ``_DIGIT_JOIN_GAP``. The first group wins: a row carries a single
        carton value, so extra far-apart groups are stray marks, not data.
        """
        ranked = sorted(words, key=_xcenter)
        groups: list[list[dict]] = [[ranked[0]]]
        for w in ranked[1:]:
            prev = groups[-1][-1]
            if float(w["x0"]) - float(prev["x1"]) <= _DIGIT_JOIN_GAP:
                groups[-1].append(w)
            else:
                groups.append([w])
        return "".join(w["text"].strip() for w in groups[0])

    @staticmethod
    def _carton_value(text: str) -> RawValue:
        """Wrap a carton count verbatim, parsing it as an int when whole."""
        cleaned = text.replace(",", "")
        try:
            number = float(cleaned)
        except ValueError:
            return RawValue(raw_text=text, parsed=text)
        parsed: float | int = int(number) if number.is_integer() else number
        return RawValue(raw_text=text, parsed=parsed)
