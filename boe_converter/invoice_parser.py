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

import pdfplumber

from boe_converter.models import RawValue

# Units seen in the invoice's quantity column; used to recognize a data row and
# to bound the carton column on its right.
_UNIT_RE = re.compile(r"^[A-Za-z]{2,4}$")
_UNITS = {"PCS", "DOZ", "KGS", "GRS", "SET", "NOS", "UNT", "MTR", "PRS", "BOX"}

# Horizontal half-width (points) of the carton column band around the ``CTNS``
# header token's centre. The invoice's CTNS values sit within ~15pt of the
# header centre; the quantity column is ~55pt to the right, well outside.
_CTNS_BAND = 22.0

# A row's SR NO sits in the far-left column.
_SR_NO_MAX_X = 70.0


class InvoicePackingListParser:
    """Extracts per-line carton counts from an invoice / packing-list PDF."""

    ROW_TOLERANCE = 3.0

    def parse_cartons(self, doc) -> dict[int, RawValue]:
        """Return ``{serial: cartons}`` read from the invoice pages of ``doc``.

        ``doc`` may be a path, bytes buffer, or an opened ``pdfplumber.PDF``.
        Only the invoice line table is read (the packing-list pages and the
        totals row are skipped); blank carton cells produce no entry.
        """
        handle, pages, should_close = self._resolve(doc)
        try:
            cartons: dict[int, RawValue] = {}
            for page in pages:
                text = page.extract_text() or ""
                if not self._is_invoice_page(text):
                    continue
                self._parse_page(page, cartons)
            return cartons
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
        return (float(w["x0"]) + float(w["x1"])) / 2.0

    def _ctns_center(self, rows: list[list[dict]]) -> float | None:
        """Locate the ``CTNS`` header token's horizontal centre on the page."""
        for row in rows:
            for w in row:
                if w["text"].strip().upper() == "CTNS":
                    return self._center(w)
        return None

    def _parse_page(self, page, cartons: dict[int, RawValue]) -> None:
        rows = self._rows(page)
        ctns_center = self._ctns_center(rows)
        if ctns_center is None:
            return
        lo, hi = ctns_center - _CTNS_BAND, ctns_center + _CTNS_BAND

        for row in rows:
            ordered = sorted(row, key=lambda w: float(w["x0"]))
            # SR NO: an integer in the far-left column.
            sr_word = next(
                (w for w in ordered
                 if self._center(w) < _SR_NO_MAX_X and w["text"].strip().isdigit()),
                None,
            )
            if sr_word is None:
                continue
            # Require a unit token (e.g. PCS/DOZ) so totals/footer rows are
            # ignored - they have a serial-like number but no unit.
            has_unit = any(
                _UNIT_RE.match(w["text"].strip())
                and w["text"].strip().upper() in _UNITS
                for w in ordered
            )
            if not has_unit:
                continue

            serial = int(sr_word["text"].strip())
            # Carton count: a numeric token whose centre falls in the CTNS band.
            ctn_word = next(
                (w for w in ordered
                 if lo <= self._center(w) <= hi
                 and self._is_number(w["text"])),
                None,
            )
            if ctn_word is None:
                continue  # blank carton cell -> leave this line blank
            cartons[serial] = self._carton_value(ctn_word["text"].strip())

    @staticmethod
    def _is_number(text: str) -> bool:
        return bool(re.fullmatch(r"\d+(?:\.\d+)?", text.strip().replace(",", "")))

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
