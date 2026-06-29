"""Shared, immutable data models for the BOE -> CTN Excel converter.

These frozen dataclasses are the contracts that thread through every component
(parser -> calculator -> excel writer -> orchestrator). They are defined here so
all components share one authoritative definition.

Design references:
- "Data Models" section of design.md (Extracted, Computed, and Anomaly models).
- The guiding constraint is parsing fidelity: values are captured verbatim
  (``RawValue.raw_text``) and never silently dropped or substituted.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

# ---------------------------------------------------------------------------
# Extracted (verbatim) models
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class RawValue:
    """A value as printed in the BOE plus its parsed interpretation.

    The raw printed text is always preserved so nothing is silently reformatted
    or lost. ``parsed`` carries the non-destructive numeric/text interpretation.
    """

    raw_text: str | None = None          # exactly as printed; None if field absent
    parsed: float | str | None = None    # numeric/text interpretation; None if unparseable
    is_missing: bool = False             # field could not be located at all (Req 2.10/3.15)
    is_unparseable: bool = False         # located but characters not resolvable (Req 9.2/9.3)

    @classmethod
    def missing(cls) -> "RawValue":
        """Build a RawValue representing a field that could not be located."""
        return cls(raw_text=None, parsed=None, is_missing=True, is_unparseable=False)

    @classmethod
    def unparseable(cls, raw_text: str) -> "RawValue":
        """Build a RawValue for text that was found but could not be parsed."""
        return cls(raw_text=raw_text, parsed=None, is_missing=False, is_unparseable=True)


@dataclass(frozen=True)
class HeaderBlock:
    """Document-level fields for the Excel Header_Block (rows 1-8)."""

    company_name: str             # configuration value (e.g. "Gemini Unicom LLP")
    party_name: RawValue          # supplier/third-party name (Req 2.7)
    usd_rate: float               # User-supplied (Open Q5)
    details: RawValue             # e.g. "CO-32 CTN-1357" (Req 4.6)
    invoice_no: RawValue          # Req 2.3
    invoice_date: RawValue        # Req 2.4
    be_no: RawValue               # Req 2.1
    be_date: RawValue             # Req 2.2
    bl_no: RawValue               # Req 2.9 (WHERE present)
    bl_date: RawValue             # Req 2.9
    invoice_amount: RawValue      # Req 2.5
    invoice_currency: RawValue    # Req 2.5
    package_count: RawValue       # PKG/CTN total e.g. 1357 (Req 2.6, 7.5)
    container_details: RawValue   # container number + count (Req 2.8)


@dataclass(frozen=True)
class LineItem:
    """One BOE line item, joined from Part II (invoice) and Part III (duty)."""

    item_serial: int              # Req 3.4 (join key)
    cth_hsn: RawValue             # Req 3.5
    description: RawValue         # verbatim, untruncated (Req 3.6)
    unit_price_usd: RawValue      # UPI (Req 3.7)
    quantity: RawValue            # Req 3.8
    unit: RawValue                # UQC (Req 3.8)
    assessable_value: RawValue    # Req 3.9
    bcd_rate: RawValue            # Req 3.10
    bcd_amount: RawValue          # Req 3.10
    igst_rate: RawValue           # Req 3.11
    total_duty: RawValue          # Req 3.12


@dataclass(frozen=True)
class ExtractedDocument:
    """The full verbatim extraction result for a BOE PDF."""

    header: HeaderBlock
    line_items: list[LineItem]            # one per BOE item, ascending serial
    declared_item_count: int | None = None
    flags: list["ReviewFlag"] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Computed models
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ComputedLine:
    """Derived per-line monetary values (full precision, no rounding).

    Any value depending on a missing/non-numeric input is left ``None`` (written
    blank) and the line is flagged for review (Req 6.13).
    """

    source: LineItem
    amount_usd: float | None = None              # Req 6.1  = unit_price_usd * qty
    purchase_inr: float | None = None            # Req 6.2  = amount_usd * usd_rate
    sws_amount: float | None = None              # Req 6.3  = bcd_amount * 0.10
    total_customs_duty: float | None = None      # Req 6.4  = bcd_amount + sws_amount
    igst_amount: float | None = None             # Req 6.5  = igst_rate * (assess + total_customs_duty)
    combined_duty: float | None = None           # Req 6.6  = total_customs_duty + igst_amount
    land_cost_excl_gst: float | None = None      # Req 6.7  = purchase_inr + total_customs_duty
    land_cost_incl_gst: float | None = None      # Req 6.8  = land_cost_excl_gst + igst_amount
    pcs: float | None = None                     # Req 5.8  = qty*12 when unit=="DOZ"
    purchase_rate_per_unit: float | None = None  # Req 6.9  = land_cost_excl_gst / qty ; 0 when qty==0


@dataclass(frozen=True)
class Totals:
    """Column-wise sums for the Totals_Row (row 61). All zero when no lines."""

    total_amount_usd: float = 0.0        # Req 7.1
    total_assessable_value: float = 0.0  # Req 7.2
    total_customs_duty: float = 0.0      # Req 7.2
    total_igst: float = 0.0              # Req 7.2
    total_land_cost_excl_gst: float = 0.0
    total_land_cost_incl_gst: float = 0.0
    package_count: RawValue = field(default_factory=RawValue)  # Req 7.5


@dataclass(frozen=True)
class ComputedDocument:
    """The full computed result ready for the Excel_Generator."""

    header: HeaderBlock
    lines: list[ComputedLine]
    totals: Totals
    flags: list["ReviewFlag"] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Anomaly model
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ReviewFlag:
    """A single anomaly: a missing/unparseable field or a recompute mismatch."""

    scope: Literal["header", "line_item", "totals"]
    field_name: str                                # e.g. "assessable_value"
    reason: Literal["MISSING", "UNPARSEABLE", "RECOMPUTE_MISMATCH"]
    item_serial: int | None = None                 # set for line_item scope
    raw_text: str | None = None                    # preserved raw text for UNPARSEABLE (Req 9.2)


@dataclass(frozen=True)
class Discrepancy:
    """A cross-check mismatch surfaced to the User (never a silent drop)."""

    kind: Literal["ITEM_COUNT", "INVOICE_TOTAL", "RECOMPUTE"]
    message: str
    expected: float | int | str | None = None   # declared/extracted value
    actual: float | int | str | None = None     # computed/recomputed value


# ---------------------------------------------------------------------------
# Summary model (UI)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ConversionSummary:
    """The conversion outcome rendered in the Web_Interface (Req 9.1)."""

    line_items_extracted: int                       # Req 9.1
    declared_item_count: int | None                 # Req 3.2/3.3
    total_invoice_amount_usd: float                 # Req 9.1
    declared_invoice_amount_usd: float | None
    review_flag_count: int                          # Req 9.1
    review_flags: list[ReviewFlag] = field(default_factory=list)
    discrepancies: list[Discrepancy] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Helper: ReviewFlagSet
# ---------------------------------------------------------------------------


class ReviewFlagSet:
    """A mutable collection helper for aggregating ``ReviewFlag``s.

    Threads through extraction, computation, and writing so components can
    accumulate flags and the Excel_Generator can quickly answer "is this
    header field / line-item field flagged?" to decide whether to blank a cell.
    """

    def __init__(self, flags: list[ReviewFlag] | None = None) -> None:
        self._flags: list[ReviewFlag] = list(flags) if flags else []

    # -- mutation -----------------------------------------------------------
    def add(self, flag: ReviewFlag) -> None:
        """Add a single review flag."""
        self._flags.append(flag)

    def extend(self, flags: list[ReviewFlag]) -> None:
        """Add many review flags."""
        self._flags.extend(flags)

    # -- access -------------------------------------------------------------
    @property
    def flags(self) -> list[ReviewFlag]:
        """Return a copy of the accumulated flags."""
        return list(self._flags)

    def __len__(self) -> int:
        return len(self._flags)

    def __iter__(self):
        return iter(self._flags)

    def count(self) -> int:
        """Total number of flags (Req 9.1 review_flag_count)."""
        return len(self._flags)

    # -- queries ------------------------------------------------------------
    def is_header_flagged(self, field_name: str) -> bool:
        """True if the named header field has a review flag."""
        return any(
            f.scope == "header" and f.field_name == field_name for f in self._flags
        )

    def is_line_flagged(self, item_serial: int, field_name: str) -> bool:
        """True if the given line-item field has a review flag."""
        return any(
            f.scope == "line_item"
            and f.item_serial == item_serial
            and f.field_name == field_name
            for f in self._flags
        )

    def for_line(self, item_serial: int) -> list[ReviewFlag]:
        """All flags raised against a given line item."""
        return [
            f
            for f in self._flags
            if f.scope == "line_item" and f.item_serial == item_serial
        ]
