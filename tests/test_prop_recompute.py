"""Property test for the extracted-vs-recomputed mismatch cross-check.

**Property 14: Extracted-vs-recomputed mismatches beyond tolerance are reported**

*For any* numeric value extracted from the BOE that differs from the value
recomputed from its related fields by more than 0.01 in that value's unit, a
``Discrepancy`` of kind ``RECOMPUTE`` is produced reporting both the extracted
and the recomputed value; differences within 0.01 produce none.

**Validates: Requirements 9.4**

The orchestrator's recompute check (``_check_recompute``) compares each line's
*extracted* ``total_duty`` against the calculator's *recomputed*
``combined_duty`` (= ``(BCD + SWS) + IGST``). This test drives the real
``ValueCalculator`` through the ``ConversionOrchestrator`` with fakes for the
validator, parser, and Excel generator so the only thing under test is the
declared-vs-recomputed RECOMPUTE reporting.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from hypothesis import given
from hypothesis import strategies as st

from boe_converter.calculator import ValueCalculator
from boe_converter.models import (
    ExtractedDocument,
    HeaderBlock,
    LineItem,
    RawValue,
    ReviewFlagSet,
)
from boe_converter.orchestrator import TOLERANCE, ConversionOrchestrator

# A line number that begins the per-line RECOMPUTE message (design wording:
# "Line item {serial}: extracted total duty ...").
_SERIAL_RE = re.compile(r"Line item (\d+):")

# Bounded, finite numeric inputs keep the recomputed combined_duty in a range
# where an offset of >= 0.02 is always clearly beyond the 0.01 tolerance and an
# offset of exactly 0 is always within it (no float-rounding boundary games).
_money = st.floats(min_value=0.0, max_value=1_000_000.0, allow_nan=False, allow_infinity=False)
_rate = st.floats(min_value=0.0, max_value=2.0, allow_nan=False, allow_infinity=False)
# A discrepancy-inducing offset: magnitude strictly and comfortably > TOLERANCE.
_offset = st.floats(min_value=0.02, max_value=100_000.0, allow_nan=False, allow_infinity=False)


@dataclass(frozen=True)
class _LineSpec:
    """The generated inputs for one line plus whether it should mismatch."""

    bcd_amount: float
    igst_rate: float
    assessable_value: float
    unit_price: float
    quantity: float
    mismatch: bool
    offset: float  # signed; only applied when ``mismatch`` is True


@st.composite
def _line_specs(draw) -> list[_LineSpec]:
    """Generate 1..8 line specs, each flagged to match or to mismatch."""

    count = draw(st.integers(min_value=1, max_value=8))
    specs: list[_LineSpec] = []
    for _ in range(count):
        mismatch = draw(st.booleans())
        magnitude = draw(_offset)
        sign = draw(st.sampled_from((1.0, -1.0)))
        specs.append(
            _LineSpec(
                bcd_amount=draw(_money),
                igst_rate=draw(_rate),
                assessable_value=draw(_money),
                unit_price=draw(_money),
                quantity=draw(_money),
                mismatch=mismatch,
                offset=sign * magnitude,
            )
        )
    return specs


def _num(value: float) -> RawValue:
    """A RawValue carrying a parseable numeric interpretation."""

    return RawValue(raw_text=repr(value), parsed=value)


def _recomputed_combined_duty(spec: _LineSpec) -> float:
    """Recompute combined_duty exactly as ``ValueCalculator`` does.

    Mirrors the calculator's float operations bit-for-bit so the expected value
    equals the one the orchestrator compares against.
    """

    total_customs_duty = spec.bcd_amount + spec.bcd_amount * ValueCalculator.SWS_RATE
    igst_amount = spec.igst_rate * (spec.assessable_value + total_customs_duty)
    return total_customs_duty + igst_amount


def _build_extracted(specs: list[_LineSpec]) -> tuple[ExtractedDocument, dict[int, tuple[float, float]], set[int]]:
    """Build an ExtractedDocument from specs.

    Returns the document, a ``serial -> (extracted_duty, recomputed_duty)`` map,
    and the set of serials whose extracted total_duty is beyond tolerance.
    """

    lines: list[LineItem] = []
    by_serial: dict[int, tuple[float, float]] = {}
    expected_mismatch: set[int] = set()

    for index, spec in enumerate(specs, start=1):
        recomputed = _recomputed_combined_duty(spec)
        extracted_duty = recomputed + (spec.offset if spec.mismatch else 0.0)
        lines.append(
            LineItem(
                item_serial=index,
                cth_hsn=RawValue.missing(),
                description=RawValue.missing(),
                unit_price_usd=_num(spec.unit_price),
                quantity=_num(spec.quantity),
                unit=RawValue.missing(),
                assessable_value=_num(spec.assessable_value),
                bcd_rate=RawValue.missing(),
                bcd_amount=_num(spec.bcd_amount),
                igst_rate=_num(spec.igst_rate),
                total_duty=_num(extracted_duty),
            )
        )
        by_serial[index] = (extracted_duty, recomputed)
        if abs(extracted_duty - recomputed) > TOLERANCE:
            expected_mismatch.add(index)

    header = HeaderBlock(
        company_name="Gemini Unicom LLP",
        party_name=RawValue.missing(),
        usd_rate=80.0,
        details=RawValue.missing(),
        invoice_no=RawValue.missing(),
        invoice_date=RawValue.missing(),
        be_no=RawValue.missing(),
        be_date=RawValue.missing(),
        bl_no=RawValue.missing(),
        bl_date=RawValue.missing(),
        invoice_amount=RawValue.missing(),
        invoice_currency=RawValue.missing(),
        package_count=RawValue.missing(),
        container_details=RawValue.missing(),
    )
    extracted = ExtractedDocument(
        header=header,
        line_items=lines,
        declared_item_count=len(lines),
        flags=[],
    )
    return extracted, by_serial, expected_mismatch


class _OkValidator:
    """Fake validator: every upload passes, carrying no PDF handle."""

    def validate(self, raw, filename):
        from boe_converter.validator import ValidationOutcome

        return ValidationOutcome.success(handle=None)


class _FakeParser:
    """Fake parser: returns a pre-built ExtractedDocument, ignoring the input."""

    def __init__(self, extracted: ExtractedDocument) -> None:
        self._extracted = extracted

    def parse(self, doc, usd_rate):
        return self._extracted


class _FakeGenerator:
    """Fake Excel generator: returns fixed bytes for a successful build."""

    def generate(self, computed, flags: ReviewFlagSet) -> bytes:
        return b"xlsx"


@given(specs=_line_specs())
def test_recompute_mismatches_beyond_tolerance_are_reported(specs: list[_LineSpec]) -> None:
    """RECOMPUTE discrepancies appear exactly for out-of-tolerance lines."""

    extracted, by_serial, expected_mismatch = _build_extracted(specs)

    orchestrator = ConversionOrchestrator(
        validator=_OkValidator(),
        parser=_FakeParser(extracted),
        calculator=ValueCalculator(),
        generator=_FakeGenerator(),
    )

    result = orchestrator.convert(b"pdf-bytes", "boe.pdf", usd_rate=80.0)

    assert result.ok is True
    assert result.summary is not None

    recompute_discs = [
        d for d in result.summary.discrepancies if d.kind == "RECOMPUTE"
    ]

    # Exactly one RECOMPUTE discrepancy per out-of-tolerance line; none for the
    # in-tolerance (matching) lines.
    reported_serials: set[int] = set()
    for disc in recompute_discs:
        match = _SERIAL_RE.search(disc.message)
        assert match is not None, f"message lacks a serial: {disc.message!r}"
        serial = int(match.group(1))
        reported_serials.add(serial)

        extracted_duty, recomputed_duty = by_serial[serial]
        # The discrepancy carries both the extracted and the recomputed value.
        assert disc.expected == extracted_duty
        assert disc.actual == recomputed_duty

    assert len(recompute_discs) == len(reported_serials)  # no duplicate serials
    assert reported_serials == expected_mismatch
