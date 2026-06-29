"""Value_Calculator: pure, deterministic per-line and totals computation.

Implements the normative per-line monetary formulas from Requirement 6 (and the
``pcs`` rule from Requirement 5.8/5.9) at full floating-point precision. Any
required input that is missing or non-numeric leaves every dependent value
``None`` (written blank downstream) and produces a ``ReviewFlag`` for the
affected line item (Req 6.13); no default value is ever substituted.

Normative formulas (design.md -> Computation model):

    amount_usd             = unit_price * qty
    purchase_inr           = amount_usd * usd_rate
    sws_amount             = bcd_amount * 0.10
    total_customs_duty     = bcd_amount + sws_amount
    igst_amount            = igst_rate * (assessable_value + total_customs_duty)
    combined_duty          = total_customs_duty + igst_amount
    land_cost_excl_gst     = purchase_inr + total_customs_duty
    land_cost_incl_gst     = land_cost_excl_gst + igst_amount
    purchase_rate_per_unit = land_cost_excl_gst / qty   (= 0 when qty == 0)
    pcs                    = qty * 12   (only when unit trimmed/upper == "DOZ")
"""

from __future__ import annotations

from boe_converter.models import (
    ComputedDocument,
    ComputedLine,
    ExtractedDocument,
    LineItem,
    RawValue,
    ReviewFlag,
    Totals,
)

# Required computation inputs (Req 6.13). USD rate is supplied per-conversion as
# an argument; the rest are per-line extracted fields on ``LineItem``.
_REQUIRED_LINE_FIELDS = (
    "unit_price_usd",
    "quantity",
    "assessable_value",
    "bcd_amount",
    "igst_rate",
)


def _coerce_number(value: object) -> float | None:
    """Return ``value`` as a float, or ``None`` if it is not a numeric value.

    Booleans are explicitly rejected (``bool`` is a subclass of ``int`` but is
    not a meaningful monetary input). Numeric strings are tolerantly coerced so
    a value already resolved to text but holding digits is still usable; any
    non-numeric string yields ``None`` (treated as non-numeric per Req 6.13).
    """

    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def _as_number(rv: RawValue | None) -> float | None:
    """Extract the numeric value of a ``RawValue`` (or ``None`` if unusable)."""

    if rv is None or rv.is_missing or rv.is_unparseable:
        return None
    return _coerce_number(rv.parsed)


def _unit_text(rv: RawValue | None) -> str | None:
    """Return the unit's textual content, preferring the parsed interpretation."""

    if rv is None:
        return None
    if isinstance(rv.parsed, str):
        return rv.parsed
    return rv.raw_text


def _sum_optional(values) -> float:
    """Sum an iterable of ``float | None`` values, skipping ``None``.

    ``None`` represents a per-line value left blank because a required input was
    missing/non-numeric (Req 6.13); it contributes nothing to the column total.
    An empty (or all-``None``) iterable yields ``0.0`` (Req 7.3). The result is
    kept at full floating-point precision with no rounding.
    """

    total = 0.0
    for value in values:
        if value is not None:
            total += value
    return total


def _as_raw_value(pkg_count: RawValue | int | float | str | None) -> RawValue:
    """Carry the package count through as a ``RawValue`` (Req 7.5).

    If it is already a ``RawValue`` it is passed through unchanged. ``None``
    becomes a missing ``RawValue``; a numeric/text value is wrapped, preserving
    its printed form in ``raw_text`` and a non-destructive numeric parse.
    """

    if isinstance(pkg_count, RawValue):
        return pkg_count
    if pkg_count is None:
        return RawValue.missing()
    parsed = _coerce_number(pkg_count)
    return RawValue(raw_text=str(pkg_count), parsed=parsed if parsed is not None else pkg_count)


class ValueCalculator:
    """Pure function: same inputs always produce the same outputs, no I/O."""

    SWS_RATE = 0.10  # Social Welfare Surcharge fixed at 10% of BCD (Req 6.3)

    def compute_line(self, item: LineItem, usd_rate: float) -> ComputedLine:
        """Compute per-line derived values per Requirement 6 / 5.8.

        Returns the ``ComputedLine``; any value whose required input is missing
        or non-numeric is left ``None`` at full floating-point precision (no
        rounding). Use :meth:`line_review_flags` to obtain the review flags for
        the same inputs (aggregated by ``compute`` in task 2.5).
        """

        line, _flags = self._compute_line(item, usd_rate)
        return line

    def line_review_flags(self, item: LineItem, usd_rate: float) -> list[ReviewFlag]:
        """Return the review flags raised for a line's missing/non-numeric inputs."""

        _line, flags = self._compute_line(item, usd_rate)
        return flags

    # -- internal -----------------------------------------------------------
    def _compute_line(
        self, item: LineItem, usd_rate: float
    ) -> tuple[ComputedLine, list[ReviewFlag]]:
        """Do the actual per-line computation, returning the line and its flags."""

        # Resolve required numeric inputs (None => missing or non-numeric).
        unit_price = _as_number(item.unit_price_usd)
        qty = _as_number(item.quantity)
        assessable = _as_number(item.assessable_value)
        bcd_amount = _as_number(item.bcd_amount)
        igst_rate = _as_number(item.igst_rate)
        rate = _coerce_number(usd_rate)

        flags = self._collect_flags(
            item,
            unit_price=unit_price,
            qty=qty,
            assessable=assessable,
            bcd_amount=bcd_amount,
            igst_rate=igst_rate,
            rate=rate,
        )

        # Derived values, propagating None when any required input is absent.
        amount_usd = (
            unit_price * qty if unit_price is not None and qty is not None else None
        )
        purchase_inr = (
            amount_usd * rate if amount_usd is not None and rate is not None else None
        )
        sws_amount = bcd_amount * self.SWS_RATE if bcd_amount is not None else None
        total_customs_duty = (
            bcd_amount + sws_amount
            if bcd_amount is not None and sws_amount is not None
            else None
        )
        igst_amount = (
            igst_rate * (assessable + total_customs_duty)
            if igst_rate is not None
            and assessable is not None
            and total_customs_duty is not None
            else None
        )
        combined_duty = (
            total_customs_duty + igst_amount
            if total_customs_duty is not None and igst_amount is not None
            else None
        )
        land_cost_excl_gst = (
            purchase_inr + total_customs_duty
            if purchase_inr is not None and total_customs_duty is not None
            else None
        )
        land_cost_incl_gst = (
            land_cost_excl_gst + igst_amount
            if land_cost_excl_gst is not None and igst_amount is not None
            else None
        )
        purchase_rate_per_unit = self._purchase_rate_per_unit(qty, land_cost_excl_gst)
        pcs = self._pcs(item.unit, qty)

        line = ComputedLine(
            source=item,
            amount_usd=amount_usd,
            purchase_inr=purchase_inr,
            sws_amount=sws_amount,
            total_customs_duty=total_customs_duty,
            igst_amount=igst_amount,
            combined_duty=combined_duty,
            land_cost_excl_gst=land_cost_excl_gst,
            land_cost_incl_gst=land_cost_incl_gst,
            pcs=pcs,
            purchase_rate_per_unit=purchase_rate_per_unit,
        )
        return line, flags

    @staticmethod
    def _purchase_rate_per_unit(
        qty: float | None, land_cost_excl_gst: float | None
    ) -> float | None:
        """Land cost (excl GST) per unit; 0 when qty == 0 (Req 6.9/6.10)."""

        if qty is None:
            return None
        if qty == 0:
            return 0  # avoid division by zero (Req 6.10)
        if land_cost_excl_gst is None:
            return None
        return land_cost_excl_gst / qty

    @staticmethod
    def _pcs(unit: RawValue | None, qty: float | None) -> float | None:
        """qty * 12 only when the unit trimmed/upper-cased == "DOZ" (Req 5.8/5.9)."""

        text = _unit_text(unit)
        if text is None or text.strip().upper() != "DOZ":
            return None
        if qty is None:
            return None
        return qty * 12

    def _collect_flags(
        self,
        item: LineItem,
        *,
        unit_price: float | None,
        qty: float | None,
        assessable: float | None,
        bcd_amount: float | None,
        igst_rate: float | None,
        rate: float | None,
    ) -> list[ReviewFlag]:
        """Build a ReviewFlag for each missing/non-numeric required input (Req 6.13)."""

        flags: list[ReviewFlag] = []
        resolved = {
            "unit_price_usd": unit_price,
            "quantity": qty,
            "assessable_value": assessable,
            "bcd_amount": bcd_amount,
            "igst_rate": igst_rate,
        }
        for field_name in _REQUIRED_LINE_FIELDS:
            if resolved[field_name] is None:
                flags.append(self._line_flag(item, field_name, getattr(item, field_name)))
        if rate is None:
            # USD rate is supplied per-conversion, so there is no RawValue.
            flags.append(
                ReviewFlag(
                    scope="line_item",
                    field_name="usd_rate",
                    reason="MISSING",
                    item_serial=item.item_serial,
                )
            )
        return flags

    @staticmethod
    def _line_flag(item: LineItem, field_name: str, rv: RawValue | None) -> ReviewFlag:
        """Create a line-item ReviewFlag, choosing MISSING vs UNPARSEABLE."""

        if rv is not None and not rv.is_missing:
            reason = "UNPARSEABLE"
        else:
            reason = "MISSING"
        return ReviewFlag(
            scope="line_item",
            field_name=field_name,
            reason=reason,
            item_serial=item.item_serial,
            raw_text=rv.raw_text if rv is not None else None,
        )

    def compute_totals(
        self, lines: list[ComputedLine], pkg_count: RawValue | int | float | str | None
    ) -> Totals:
        """Compute the column-wise sums for the Totals_Row (Req 7.1, 7.2, 7.3, 7.5).

        Each total is the sum of the corresponding per-line values; per-line
        values that are ``None`` (a missing/non-numeric input left blank per Req
        6.13) do not contribute to the sum. When ``lines`` is empty every total
        is ``0`` (Req 7.3). Sums are kept at full floating-point precision (no
        rounding). The BOE package count is carried through unchanged as a
        ``RawValue`` (Req 7.5).

        Column mapping (design.md totals cell map): L -> amount_usd,
        N -> assessable_value (direct extracted), O -> land_cost_excl_gst,
        P -> total_customs_duty, Q -> igst_amount, X -> land_cost_incl_gst.
        """

        total_amount_usd = _sum_optional(line.amount_usd for line in lines)
        total_assessable_value = _sum_optional(
            _as_number(line.source.assessable_value) for line in lines
        )
        total_customs_duty = _sum_optional(line.total_customs_duty for line in lines)
        total_igst = _sum_optional(line.igst_amount for line in lines)
        total_land_cost_excl_gst = _sum_optional(
            line.land_cost_excl_gst for line in lines
        )
        total_land_cost_incl_gst = _sum_optional(
            line.land_cost_incl_gst for line in lines
        )

        return Totals(
            total_amount_usd=total_amount_usd,
            total_assessable_value=total_assessable_value,
            total_customs_duty=total_customs_duty,
            total_igst=total_igst,
            total_land_cost_excl_gst=total_land_cost_excl_gst,
            total_land_cost_incl_gst=total_land_cost_incl_gst,
            package_count=_as_raw_value(pkg_count),
        )

    def compute(self, doc: ExtractedDocument, usd_rate: float) -> ComputedDocument:
        """Compute the full ``ComputedDocument`` from an ``ExtractedDocument``.

        Computes every per-line value (:meth:`compute_line`), aggregates the
        per-line review flags (:meth:`line_review_flags`) onto the flags already
        carried by ``doc`` (e.g. extraction flags), builds the Totals_Row
        (:meth:`compute_totals`), and assembles the result. Pure: no I/O.
        """

        computed_lines: list[ComputedLine] = []
        flags: list[ReviewFlag] = list(doc.flags)
        for item in doc.line_items:
            line, line_flags = self._compute_line(item, usd_rate)
            computed_lines.append(line)
            flags.extend(line_flags)

        totals = self.compute_totals(computed_lines, doc.header.package_count)

        return ComputedDocument(
            header=doc.header,
            lines=computed_lines,
            totals=totals,
            flags=flags,
        )
