"""OldFormatParser: extraction for the legacy "Indian Customs EDI System" BOE.

The project handles two distinct Bill of Entry layouts:

- the **new** ICEGATE format (positional, multi-column tables with rotated margin
  labels) handled by :class:`boe_converter.parser.PdfParser`; and
- the **old** "Indian Customs EDI System - Imports V1.5R001" format handled here.

The old format is a plain, line-oriented printout with **one line item per page**
laid out as a fixed stacked block::

    slno RITC Description RSP Load PROV
    Qty Unit Price CTH C.Notn C.NSNO Cus Dty Rt BCD amt(Rs.)
    Unit Ass Val CETH E.Notn E.NSNO Exc Dty Rt CVD amt(Rs.)
    -----
    35 39199010 STATIONARY STICKER(O/T REPUTED BRAND)   <- slno, RITC, description
    298.00 0.360000 39199010 10.00 % 1033.90            <- qty, price, CTH, BCD rate, BCD amt
    Cus AIDC 011/2021 17 0.00% 0.00
    GRS 10338.80 NOEXCISE 0.00 % 0.00                   <- unit, assessable value
    ... cess lines ...
    Social Welfare Surcharge: 10.00 % 103.40            <- SWS rate, SWS amount
    IGST 009/2025 II120 18.00 % 2065.70                 <- IGST rate, IGST amount
    GST Cess 001/2017 56 0.00 % 0.00
    Rs. <cum> Page Total Rs. <cum-duty>

Because the layout is line-based (no rotated text), extraction is done with
anchored regexes over ``page.extract_text()`` rather than the geometric
row-reconstruction the new format requires. The result is the same
:class:`~boe_converter.models.ExtractedDocument` contract consumed by the
calculator and Excel generator, so the rest of the pipeline is format-agnostic.

Rate fields (BCD %, IGST %) are captured *with* their percent sign so the shared
``PdfParser._capture`` parses them to a decimal fraction (e.g. ``"18.00%"`` ->
``0.18``), matching the convention the calculator expects (``igst_amount =
igst_rate * (assessable + total_customs_duty)``). Values are preserved verbatim
in ``RawValue.raw_text``; nothing is silently dropped or reformatted.
"""

from __future__ import annotations

import re

from boe_converter.models import (
    ExtractedDocument,
    HeaderBlock,
    LineItem,
    RawValue,
    ReviewFlag,
)
from boe_converter.parser import PdfParser

# A marker present on every page of the legacy format; used for detection.
_OLD_FORMAT_MARKERS = ("Indian Customs EDI System", "V1.5R001")

# --- Header field patterns (anchored on the legacy labels) -----------------
_RE_BE = re.compile(r"No/Dt\./cc/Typ:\s*(\d+)\s*/\s*(\d{1,2}/\d{1,2}/\d{2,4})")
_RE_INV = re.compile(
    r"Inv\s+No\s+&\s+Dt\.\s*:\s*(\S+)\s+(\d{1,2}/\d{1,2}/\d{2,4})"
)
_RE_INV_VAL = re.compile(r"Inv\s+Val\s*:\s*([\d,]+\.?\d*)\s+([A-Za-z]{3})")
_RE_PKGS = re.compile(r"No\.\s*Of\s*Pkgs\.\s*:\s*(\d+)")
_RE_BL = re.compile(r"(?:^|[^/])BL\s+No\s*:\s*([A-Za-z0-9]+)", re.MULTILINE)

# --- Line-item block patterns ----------------------------------------------
# Item id line: small integer serial, 6-10 digit RITC, then the description
# (which may itself begin with a digit, e.g. "9v BATTERY"). The qty line below
# never matches this because its first token carries a decimal point.
_RE_ITEM_ID = re.compile(r"^(\d{1,3})\s+(\d{6,10})\s+(.+)$")
# Qty line: qty, unit price, CTH, BCD rate %, BCD amount.
_RE_QTY = re.compile(
    r"^([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+(\d{6,10})\s+([\d.]+)\s*%\s+([\d,]+\.?\d*)$"
)
# Unit / assessable-value line: a unit code (alpha), the assessable value, then a
# notification token (e.g. NOEXCISE) and a rate%. The "Cus AIDC ..." line never
# matches (its second token "AIDC" is not numeric).
_RE_UNIT = re.compile(r"^([A-Za-z]+)\s+([\d,]+\.?\d*)\s+\S+\s+[\d.]+\s*%")
_RE_SWS = re.compile(r"Social\s+Welfare\s+Surcharge:\s*([\d.]+)\s*%\s*([\d,]+\.?\d*)")
_RE_IGST = re.compile(r"\bIGST\s+\S+\s+\S+\s+([\d.]+)\s*%\s*([\d,]+\.?\d*)")

# Required header fields (Req 2.1-2.8); B/L (2.9) is "where present".
_REQUIRED_HEADER = (
    "be_no",
    "be_date",
    "invoice_no",
    "invoice_date",
    "invoice_amount",
    "invoice_currency",
    "package_count",
    "party_name",
)


def is_old_format(text: str) -> bool:
    """Return ``True`` when ``text`` (a page's text) is the legacy EDI format."""
    return any(marker in text for marker in _OLD_FORMAT_MARKERS)


class OldFormatParser:
    """Extracts an :class:`ExtractedDocument` from a legacy EDI-format BOE.

    Operates on a list of already-opened ``pdfplumber`` pages (the caller owns
    the handle), so :class:`PdfParser` can detect the format once and delegate
    without re-opening the document.
    """

    def __init__(self) -> None:
        # Reuse the shared verbatim-capture helper so RawValue conventions
        # (raw_text preserved, non-destructive numeric parse, percent -> fraction)
        # are identical across both parsers.
        self._cap = PdfParser()

    # ------------------------------------------------------------------
    def parse(
        self,
        pages,
        *,
        company_name: str = "Gemini Unicom LLP",
        usd_rate: float = 0.0,
    ) -> ExtractedDocument:
        """Parse the legacy-format ``pages`` into an ``ExtractedDocument``."""
        page_texts = [(p.extract_text() or "") for p in pages]
        full_text = "\n".join(page_texts)

        header, header_flags = self._extract_header(
            pages, page_texts, full_text, company_name=company_name, usd_rate=usd_rate
        )
        line_items, item_flags = self._extract_items(page_texts)
        declared_count = len(line_items)

        return ExtractedDocument(
            header=header,
            line_items=line_items,
            declared_item_count=declared_count,
            flags=[*header_flags, *item_flags],
        )

    # ------------------------------------------------------------------
    # Header
    # ------------------------------------------------------------------
    def _extract_header(
        self, pages, page_texts, full_text, *, company_name, usd_rate
    ) -> tuple[HeaderBlock, list[ReviewFlag]]:
        be_no = be_date = RawValue.missing()
        m = _RE_BE.search(full_text)
        if m:
            be_no = self._cap._capture(m.group(1))
            be_date = self._cap._capture(m.group(2))

        invoice_no = invoice_date = RawValue.missing()
        m = _RE_INV.search(full_text)
        if m:
            invoice_no = self._cap._capture(m.group(1))
            invoice_date = self._cap._capture(m.group(2))

        invoice_amount = invoice_currency = RawValue.missing()
        m = _RE_INV_VAL.search(full_text)
        if m:
            invoice_amount = self._cap._capture(m.group(1), numeric=True)
            invoice_currency = self._cap._capture(m.group(2))

        package_count = RawValue.missing()
        m = _RE_PKGS.search(full_text)
        if m:
            package_count = self._cap._capture(m.group(1), numeric=True)

        bl_no = RawValue.missing()
        m = _RE_BL.search(full_text)
        if m:
            bl_no = self._cap._capture(m.group(1))
        bl_date = self._extract_bl_date(full_text)

        party_name = self._extract_party_name(pages, page_texts)
        container_details = self._extract_container(full_text)

        header = HeaderBlock(
            company_name=company_name,
            party_name=party_name,
            usd_rate=usd_rate,
            details=RawValue.missing(),
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

        flags: list[ReviewFlag] = []
        for name in _REQUIRED_HEADER:
            value: RawValue = getattr(header, name)
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

    def _extract_bl_date(self, full_text: str) -> RawValue:
        """BL date: the date on the ``Date :`` line that follows ``BL No :``."""
        lines = full_text.splitlines()
        for i, ln in enumerate(lines):
            if re.search(r"(?:^|[^/])BL\s+No\s*:", ln):
                # The BL date is printed on the next "Date :" line (left column).
                for nxt in lines[i + 1 : i + 4]:
                    if nxt.strip().startswith("Date"):
                        dm = re.search(r"(\d{1,2}/\d{1,2}/\d{2,4})", nxt)
                        if dm:
                            return self._cap._capture(dm.group(1))
                        return RawValue.missing()
        return RawValue.missing()

    def _extract_party_name(self, pages, page_texts) -> RawValue:
        """Supplier/party name printed in the right column of the invoice row.

        On the first page the supplier sits to the right of the ``Inv No & Dt.``
        value and may wrap to the next visual line. The words are read
        positionally (x to the right of the invoice date) and joined.
        """
        if not pages:
            return RawValue.missing()
        try:
            words = pages[0].extract_words()
        except Exception:
            words = []
        # Locate the invoice-date token to find the right-column x threshold.
        inv_top = None
        date_x1 = None
        for w in words:
            if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{2,4}", w["text"]) and inv_top is None:
                # First date on the page belongs to the Inv No & Dt. row only if
                # an "Inv" label shares its row; verify by scanning the row.
                top = float(w["top"])
                row = [x for x in words if abs(float(x["top"]) - top) <= 3]
                if any(r["text"].startswith("Inv") for r in row):
                    inv_top = top
                    date_x1 = float(w["x1"])
                    break
        if inv_top is None or date_x1 is None:
            return RawValue.missing()

        # Right-column words on the invoice row carry the supplier name (the
        # extract_text wrapping is only a rendering artifact; positionally the
        # name sits on the invoice row to the right of the date).
        right = [
            w for w in words
            if abs(float(w["top"]) - inv_top) <= 3 and float(w["x0"]) > date_x1 + 20
        ]
        right.sort(key=lambda w: float(w["x0"]))
        # Stop at any label token that begins the next field (defensive: the
        # supplier column should contain only the name).
        tokens: list[str] = []
        for w in right:
            if w["text"].startswith("Inv") or w["text"] == "Val":
                break
            tokens.append(w["text"])
        name = " ".join(t for t in tokens if t).strip()
        if not name:
            return RawValue.missing()
        return self._cap._capture(name)

    def _extract_container(self, full_text: str) -> RawValue:
        """Container marks/numbers from the ``Marks:CONTAINER NOS.`` block."""
        m = re.search(
            r"Marks:CONTAINER\s+NOS\.\s*\.*\s*([A-Za-z0-9]+)", full_text
        )
        if not m:
            return RawValue.missing()
        prefix = m.group(1)
        nums = re.search(r"&\s*Nos\s*([0-9]+)", full_text)
        if nums:
            return self._cap._capture(f"{prefix} {nums.group(1)}")
        return self._cap._capture(prefix)

    # ------------------------------------------------------------------
    # Line items
    # ------------------------------------------------------------------
    def _extract_items(
        self, page_texts
    ) -> tuple[list[LineItem], list[ReviewFlag]]:
        """Parse one stacked line-item block per page, deduped by serial."""
        items: dict[int, LineItem] = {}
        flags: list[ReviewFlag] = []

        for text in page_texts:
            lines = [ln.strip() for ln in text.splitlines()]
            block = self._find_item_block(lines)
            if block is None:
                continue
            item, item_flags = self._build_line_item(block)
            if item.item_serial in items:
                continue  # first occurrence wins (serial is the join key)
            items[item.item_serial] = item
            flags.extend(item_flags)

        ordered = [items[k] for k in sorted(items)]
        return ordered, flags

    def _find_item_block(self, lines: list[str]) -> dict | None:
        """Locate and parse the single item block on a page's lines."""
        id_idx = None
        idm = None
        for i, ln in enumerate(lines):
            if "Description" in ln or "RITC" in ln:
                continue  # the "slno RITC Description ..." header row
            m = _RE_ITEM_ID.match(ln)
            if m:
                id_idx = i
                idm = m
                break
        if id_idx is None:
            return None

        slno = int(idm.group(1))
        ritc = idm.group(2)
        desc_parts = [idm.group(3).strip()]

        block: dict = {"slno": slno, "ritc": ritc}
        # Walk the lines after the id line, collecting fields. Lines before the
        # qty line that aren't the qty line are description continuation.
        qty_seen = False
        for ln in lines[id_idx + 1 : id_idx + 16]:
            if not qty_seen:
                mq = _RE_QTY.match(ln)
                if mq:
                    qty_seen = True
                    block["qty"] = mq.group(1)
                    block["price"] = mq.group(2)
                    block["cth"] = mq.group(3)
                    block["bcd_rt"] = mq.group(4)
                    block["bcd_amt"] = mq.group(5)
                    continue
                # not yet the qty line -> description wrapped onto this line
                if ln and not ln.startswith("-"):
                    desc_parts.append(ln)
                continue
            if "unit" not in block:
                mu = _RE_UNIT.match(ln)
                if mu and not ln.startswith("Cus"):
                    block["unit"] = mu.group(1)
                    block["ass"] = mu.group(2)
            ms = _RE_SWS.search(ln)
            if ms:
                block["sws_rt"], block["sws_amt"] = ms.group(1), ms.group(2)
            mi = _RE_IGST.search(ln)
            if mi:
                block["igst_rt"], block["igst_amt"] = mi.group(1), mi.group(2)

        block["desc"] = " ".join(desc_parts).strip()
        return block

    def _build_line_item(self, block: dict) -> tuple[LineItem, list[ReviewFlag]]:
        cap = self._cap._capture

        def num(key):
            return cap(block[key], numeric=True) if key in block else RawValue.missing()

        def rate(key):
            # Capture with a percent sign so _capture yields a decimal fraction.
            return cap(f"{block[key]}%", numeric=True) if key in block else RawValue.missing()

        def txt(key):
            return cap(block[key]) if key in block else RawValue.missing()

        item = LineItem(
            item_serial=block["slno"],
            cth_hsn=txt("ritc"),
            description=txt("desc"),
            unit_price_usd=num("price"),
            quantity=num("qty"),
            unit=txt("unit"),
            assessable_value=num("ass"),
            bcd_rate=rate("bcd_rt"),
            bcd_amount=num("bcd_amt"),
            igst_rate=rate("igst_rt"),
            total_duty=RawValue.missing(),  # not printed per line in this format
        )

        # Flag missing/unparseable extracted line fields (total_duty excluded:
        # the legacy format prints no per-line total duty, so its absence is not
        # an anomaly).
        flags: list[ReviewFlag] = []
        checked = (
            "cth_hsn",
            "description",
            "unit_price_usd",
            "quantity",
            "unit",
            "assessable_value",
            "bcd_rate",
            "bcd_amount",
            "igst_rate",
        )
        for name in checked:
            rv: RawValue = getattr(item, name)
            if rv.is_missing:
                flags.append(
                    ReviewFlag(
                        scope="line_item",
                        field_name=name,
                        reason="MISSING",
                        item_serial=item.item_serial,
                    )
                )
            elif rv.is_unparseable:
                flags.append(
                    ReviewFlag(
                        scope="line_item",
                        field_name=name,
                        reason="UNPARSEABLE",
                        item_serial=item.item_serial,
                        raw_text=rv.raw_text,
                    )
                )
        return item, flags
