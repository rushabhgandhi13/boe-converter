"""Property-based test for the invoice-total verification cross-check.

Property 13: Invoice-total verification reports out-of-tolerance differences.

**Validates: Requirements 7.6**

The orchestrator's ``_check_invoice_total`` compares the computed total invoice
amount in USD (the sum of per-line ``unit_price * quantity`` amounts, exposed as
``ComputedDocument.totals.total_amount_usd``) against the declared header invoice
amount printed in the BOE. Per Requirement 7.6:

- a difference of **more than** ``TOLERANCE`` (0.01) USD yields a
  ``Discrepancy(kind="INVOICE_TOTAL")`` carrying both the declared (``expected``)
  and computed (``actual``) values, and the workbook is still retained;
- a difference of ``TOLERANCE`` USD **or less** yields no such discrepancy;
- a declared total that is absent/unreadable yields an ``INVOICE_TOTAL``
  discrepancy with ``expected=None`` and the "could not be verified" message
  (Req 7.7).

Strategy: drive the real ``ConversionOrchestrator`` through injected fakes -- a
validator that always passes (handle ``None``), a parser that returns a
generated ``ExtractedDocument`` whose line items have known numeric
``unit_price_usd``/``quantity`` (so the computed USD total is known) and whose
header ``invoice_amount`` is set to a declared value chosen to be within
tolerance, beyond tolerance, or missing -- the real ``ValueCalculator``, and an
Excel generator that returns ``b"xlsx"``. The computed total is read back from
the summary (``total_invoice_amount_usd``) to keep the assertion independent of
floating-point reconstruction in the test.
"""

from __future__ import annotations

from hypothesis import given
from hypothesis import strategies as st

from boe_converter.models import (
    ExtractedDocument,
    HeaderBlock,
    LineItem,
    RawValue,
)
from boe_converter.orchestrator import (
    MESSAGE_TOTAL_UNVERIFIABLE,
    TOLERANCE,
    ConversionOrchestrator,
)
from boe_converter.calculator import ValueCalculator
from boe_converter.validator import ValidationOutcome


# ---------------------------------------------------------------------------
# Injected fakes
# ---------------------------------------------------------------------------


class _FakeValidator:
    """Always-OK validator carrying no handle (Req: post-recognition path)."""

    def validate(self, raw: bytes, filename: str) -> ValidationOutcome:
        return ValidationOutcome.success(handle=None)


class _FakeParser:
    """Parser that returns a pre-built ``ExtractedDocument`` verbatim."""

    def __init__(self, document: ExtractedDocument) -> None:
        self._document = document

    def parse(self, handle, usd_rate: float) -> ExtractedDocument:
        return self._document


class _FakeGenerator:
    """Excel generator stub: returns fixed bytes so a workbook is retained."""

    def generate(self, computed, flags) -> bytes:
        return b"xlsx"


def _orchestrator(document: ExtractedDocument) -> ConversionOrchestrator:
    """Wire the orchestrator with fakes + the real ``ValueCalculator``."""

    return ConversionOrchestrator(
        validator=_FakeValidator(),
        parser=_FakeParser(document),
        calculator=ValueCalculator(),
        generator=_FakeGenerator(),
    )


# ---------------------------------------------------------------------------
# Helpers / strategies
# ---------------------------------------------------------------------------


def _num(value: float) -> RawValue:
    """A located, parseable numeric field as it would arrive from the parser."""

    return RawValue(raw_text=str(value), parsed=value)


def _missing_inputs(item_serial: int, unit_price: float, qty: float) -> LineItem:
    """A line item with known price/qty; all other numeric inputs missing.

    Leaving ``total_duty`` (and the duty inputs) missing keeps the recompute
    cross-check from emitting unrelated ``RECOMPUTE`` discrepancies, so the test
    isolates the ``INVOICE_TOTAL`` behaviour under test.
    """

    return LineItem(
        item_serial=item_serial,
        cth_hsn=RawValue(raw_text="1234", parsed="1234"),
        description=RawValue(raw_text="item", parsed="item"),
        unit_price_usd=_num(unit_price),
        quantity=_num(qty),
        unit=RawValue(raw_text="PCS", parsed="PCS"),
        assessable_value=RawValue.missing(),
        bcd_rate=RawValue.missing(),
        bcd_amount=RawValue.missing(),
        igst_rate=RawValue.missing(),
        total_duty=RawValue.missing(),
    )


def _header(invoice_amount: RawValue) -> HeaderBlock:
    """A header carrying the declared ``invoice_amount`` under test."""

    return HeaderBlock(
        company_name="Gemini Unicom LLP",
        party_name=RawValue(raw_text="ACME", parsed="ACME"),
        usd_rate=80.0,
        details=RawValue.missing(),
        invoice_no=RawValue(raw_text="INV-1", parsed="INV-1"),
        invoice_date=RawValue(raw_text="2026-01-01", parsed="2026-01-01"),
        be_no=RawValue(raw_text="BE-1", parsed="BE-1"),
        be_date=RawValue(raw_text="2026-01-01", parsed="2026-01-01"),
        bl_no=RawValue.missing(),
        bl_date=RawValue.missing(),
        invoice_amount=invoice_amount,
        invoice_currency=RawValue(raw_text="USD", parsed="USD"),
        package_count=RawValue(raw_text="10", parsed=10),
        container_details=RawValue.missing(),
    )


def _document(pairs, invoice_amount: RawValue) -> ExtractedDocument:
    """Build an ``ExtractedDocument`` from (unit_price, qty) pairs."""

    line_items = [
        _missing_inputs(i + 1, up, qty) for i, (up, qty) in enumerate(pairs)
    ]
    return ExtractedDocument(
        header=_header(invoice_amount),
        line_items=line_items,
        declared_item_count=None,  # avoid ITEM_COUNT discrepancy interference
        flags=[],
    )


def _expected_total(pairs) -> float:
    """Mirror the calculator: order-preserving sum of ``unit_price * qty``."""

    total = 0.0
    for unit_price, qty in pairs:
        total += unit_price * qty
    return total


def _invoice_total_discrepancies(summary):
    """The INVOICE_TOTAL discrepancies reported in a conversion summary."""

    return [d for d in summary.discrepancies if d.kind == "INVOICE_TOTAL"]


# Moderate, finite magnitudes keep per-line amounts and their order-preserving
# sum well clear of overflow and floating-point granularity near the 0.01
# tolerance boundary.
_PRICE = st.floats(min_value=-1e4, max_value=1e4, allow_nan=False, allow_infinity=False)
_QTY = st.floats(min_value=0.0, max_value=1e4, allow_nan=False, allow_infinity=False)
_PAIRS = st.lists(st.tuples(_PRICE, _QTY), min_size=1, max_size=15)
_USD_RATE = st.floats(min_value=1.0, max_value=200.0, allow_nan=False, allow_infinity=False)


# ---------------------------------------------------------------------------
# Properties
# ---------------------------------------------------------------------------


@given(
    pairs=_PAIRS,
    usd_rate=_USD_RATE,
    # A difference safely beyond the 0.01 tolerance (either sign).
    offset=st.floats(
        min_value=0.02, max_value=1e6, allow_nan=False, allow_infinity=False
    ),
    negative=st.booleans(),
)
def test_out_of_tolerance_difference_is_reported(pairs, usd_rate, offset, negative):
    """|computed - declared| > 0.01 => INVOICE_TOTAL discrepancy with both values."""

    computed_total = _expected_total(pairs)
    delta = -offset if negative else offset
    declared = computed_total + delta

    result = _orchestrator(_document(pairs, _num(declared))).convert(
        b"pdf", "boe.pdf", usd_rate
    )

    # Workbook is retained: conversion succeeds and a download token is issued.
    assert result.ok is True
    assert result.download_token is not None

    discrepancies = _invoice_total_discrepancies(result.summary)
    assert len(discrepancies) == 1
    disc = discrepancies[0]
    # Carries both the declared (expected) and computed (actual) totals.
    assert disc.expected == declared
    assert disc.actual == result.summary.total_invoice_amount_usd
    # Sanity: the difference genuinely exceeds the tolerance.
    assert abs(disc.actual - disc.expected) > TOLERANCE


@given(
    pairs=_PAIRS,
    usd_rate=_USD_RATE,
    # A difference comfortably within the 0.01 tolerance (either sign).
    offset=st.floats(
        min_value=0.0, max_value=0.009, allow_nan=False, allow_infinity=False
    ),
    negative=st.booleans(),
)
def test_within_tolerance_difference_is_not_reported(pairs, usd_rate, offset, negative):
    """|computed - declared| <= 0.01 => no INVOICE_TOTAL discrepancy."""

    computed_total = _expected_total(pairs)
    delta = -offset if negative else offset
    declared = computed_total + delta

    result = _orchestrator(_document(pairs, _num(declared))).convert(
        b"pdf", "boe.pdf", usd_rate
    )

    assert result.ok is True
    assert _invoice_total_discrepancies(result.summary) == []


@given(
    pairs=_PAIRS,
    usd_rate=_USD_RATE,
    declared=st.sampled_from(
        [RawValue.missing(), RawValue.unparseable("???")]
    ),
)
def test_missing_declared_total_is_unverifiable(pairs, usd_rate, declared):
    """An absent/unreadable declared total => unverifiable INVOICE_TOTAL flag."""

    result = _orchestrator(_document(pairs, declared)).convert(
        b"pdf", "boe.pdf", usd_rate
    )

    assert result.ok is True
    discrepancies = _invoice_total_discrepancies(result.summary)
    assert len(discrepancies) == 1
    disc = discrepancies[0]
    assert disc.message == MESSAGE_TOTAL_UNVERIFIABLE
    assert disc.expected is None
    assert disc.actual == result.summary.total_invoice_amount_usd
