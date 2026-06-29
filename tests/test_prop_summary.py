"""Property-based test for completion-summary accuracy (task 9.7).

Property 16: Completion summary accurately reflects the conversion.

**Validates: Requirements 9.1**

For any successfully converted document, the ``ConversionSummary`` built by the
``ConversionOrchestrator`` accurately reflects the conversion:

- ``line_items_extracted`` equals the number of line items the parser produced;
- ``total_invoice_amount_usd`` equals the computed sum of per-line USD amounts;
- ``review_flag_count`` equals the number of review flags actually raised
  (the aggregate ``ComputedDocument.flags``), and ``review_flags`` lists them;
- ``declared_item_count`` / ``declared_invoice_amount_usd`` match what the
  parser produced; and
- the discrepancy list is consistent with the conversion (in particular an
  ``ITEM_COUNT`` discrepancy is present exactly when the declared count is
  present and differs from the extracted count).

The orchestrator is driven with injected fakes (a validator that always admits,
a parser that returns a Hypothesis-generated ``ExtractedDocument`` with varied
numbers of line items -- some carrying missing/unparseable required fields that
produce review flags via the *real* ``ValueCalculator`` -- and a generator that
returns fixed bytes). The expected outcome is independently recomputed with a
fresh ``ValueCalculator`` over the same document; because the calculator is a
pure, deterministic function the comparisons use exact equality.
"""

from __future__ import annotations

from hypothesis import given
from hypothesis import strategies as st

from boe_converter.calculator import ValueCalculator
from boe_converter.models import (
    ExtractedDocument,
    HeaderBlock,
    LineItem,
    RawValue,
    ReviewFlag,
)
from boe_converter.orchestrator import ConversionOrchestrator


# ---------------------------------------------------------------------------
# Injected fakes
# ---------------------------------------------------------------------------


class _FakeOutcome:
    """Minimal stand-in for ``ValidationOutcome`` on the success path."""

    ok = True
    error_code = None
    message = None
    handle = None


class _FakeValidator:
    """Always admits the upload (ok=True, no handle)."""

    def validate(self, raw: bytes, filename: str) -> _FakeOutcome:
        return _FakeOutcome()


class _FakeParser:
    """Returns a pre-generated ``ExtractedDocument`` regardless of input."""

    def __init__(self, doc: ExtractedDocument) -> None:
        self._doc = doc

    def parse(self, handle, usd_rate=None) -> ExtractedDocument:
        return self._doc


class _FakeGenerator:
    """Returns fixed workbook bytes (no real Excel built)."""

    def generate(self, computed, flags) -> bytes:
        return b"xlsx"


# ---------------------------------------------------------------------------
# Generators
# ---------------------------------------------------------------------------

finite_numbers = st.floats(
    min_value=-1e6, max_value=1e6, allow_nan=False, allow_infinity=False
)
quantities = st.floats(
    min_value=-1e5, max_value=1e5, allow_nan=False, allow_infinity=False
)
igst_rates = st.floats(
    min_value=0.0, max_value=1.0, allow_nan=False, allow_infinity=False
)
usd_rates = st.floats(
    min_value=0.0, max_value=200.0, allow_nan=False, allow_infinity=False
)


def _numeric_raw(value: float) -> RawValue:
    """A cleanly-parsing numeric RawValue, as the parser emits."""

    return RawValue(raw_text=repr(value), parsed=value)


def _required_field(value: float) -> st.SearchStrategy[RawValue]:
    """A required numeric input that is sometimes numeric, missing, or garbage.

    Missing / unparseable variants drive the real calculator to raise review
    flags, so the generated documents exercise non-zero ``review_flag_count``.
    """

    return st.one_of(
        st.just(_numeric_raw(value)),
        st.just(RawValue.missing()),
        st.just(RawValue.unparseable("??")),
    )


@st.composite
def line_items(draw):
    """Build a LineItem whose required inputs may be present, missing, or garbage."""

    serial = draw(st.integers(min_value=1, max_value=10_000))
    return LineItem(
        item_serial=serial,
        cth_hsn=RawValue(raw_text="39264049", parsed="39264049"),
        description=RawValue(raw_text="ITEM", parsed="ITEM"),
        unit_price_usd=draw(_required_field(draw(finite_numbers))),
        quantity=draw(_required_field(draw(quantities))),
        unit=RawValue(raw_text="DOZ", parsed="DOZ"),
        assessable_value=draw(_required_field(draw(finite_numbers))),
        bcd_rate=_numeric_raw(draw(finite_numbers)),
        bcd_amount=draw(_required_field(draw(finite_numbers))),
        igst_rate=draw(_required_field(draw(igst_rates))),
        total_duty=_numeric_raw(draw(finite_numbers)),
    )


def _invoice_amount() -> st.SearchStrategy[RawValue]:
    """Header invoice amount: a number, or absent/unreadable."""

    return st.one_of(
        finite_numbers.map(_numeric_raw),
        st.just(RawValue.missing()),
        st.just(RawValue.unparseable("N/A")),
    )


def _header(draw) -> HeaderBlock:
    missing = RawValue.missing()
    return HeaderBlock(
        company_name="Gemini Unicom LLP",
        party_name=missing,
        usd_rate=draw(usd_rates),
        details=missing,
        invoice_no=missing,
        invoice_date=missing,
        be_no=missing,
        be_date=missing,
        bl_no=missing,
        bl_date=missing,
        invoice_amount=draw(_invoice_amount()),
        invoice_currency=RawValue(raw_text="USD", parsed="USD"),
        package_count=missing,
        container_details=missing,
    )


_header_flag_fields = st.sampled_from(["party_name", "invoice_no", "be_no"])
extraction_flags_strategy = st.lists(
    _header_flag_fields.map(
        lambda name: ReviewFlag(scope="header", field_name=name, reason="MISSING")
    ),
    max_size=2,
)


@st.composite
def extracted_documents(draw):
    """An ExtractedDocument with varied items, declared count, and flags."""

    items = draw(st.lists(line_items(), min_size=0, max_size=6))
    actual = len(items)
    # declared_item_count: absent, matching, or deliberately mismatched.
    declared = draw(
        st.one_of(
            st.none(),
            st.just(actual),
            st.integers(min_value=0, max_value=10),
        )
    )
    extraction_flags = draw(extraction_flags_strategy)
    header = _header(draw)
    return ExtractedDocument(
        header=header,
        line_items=items,
        declared_item_count=declared,
        flags=extraction_flags,
    )


# ---------------------------------------------------------------------------
# Expected-value helpers (independent of the orchestrator's internals)
# ---------------------------------------------------------------------------


def _expected_declared_amount(invoice_amount: RawValue) -> float | None:
    """The numeric interpretation the summary should report for the declared total."""

    if invoice_amount.is_missing or invoice_amount.is_unparseable:
        return None
    parsed = invoice_amount.parsed
    if isinstance(parsed, bool):
        return None
    if isinstance(parsed, (int, float)):
        return float(parsed)
    return None


# ---------------------------------------------------------------------------
# Property 16
# ---------------------------------------------------------------------------


@given(doc=extracted_documents(), usd_rate=usd_rates)
def test_summary_accurately_reflects_conversion(
    doc: ExtractedDocument, usd_rate: float
):
    """ConversionSummary mirrors the items, total, flags, and discrepancies."""

    orchestrator = ConversionOrchestrator(
        validator=_FakeValidator(),
        parser=_FakeParser(doc),
        calculator=ValueCalculator(),
        generator=_FakeGenerator(),
    )

    result = orchestrator.convert(b"%PDF-fake", "boe.pdf", usd_rate)

    # The conversion must succeed for the summary to exist (Req 9.1).
    assert result.ok is True
    assert result.download_token is not None
    summary = result.summary
    assert summary is not None

    # Independently recompute the expected computation with a fresh calculator;
    # the calculator is pure/deterministic, so values are bit-identical.
    expected = ValueCalculator().compute(doc, usd_rate)

    # line_items_extracted == number of parsed line items.
    assert summary.line_items_extracted == len(doc.line_items)

    # total_invoice_amount_usd == computed total of per-line USD amounts.
    assert summary.total_invoice_amount_usd == expected.totals.total_amount_usd

    # review_flag_count == number of review flags actually produced, and the
    # listed flags are exactly those flags.
    assert summary.review_flag_count == len(expected.flags)
    assert summary.review_flags == expected.flags
    assert len(summary.review_flags) == summary.review_flag_count

    # Declared counts / amounts match what the parser produced.
    assert summary.declared_item_count == doc.declared_item_count
    assert summary.declared_invoice_amount_usd == _expected_declared_amount(
        doc.header.invoice_amount
    )

    # Discrepancy consistency: an ITEM_COUNT discrepancy is present exactly when
    # a declared count exists and differs from the extracted count (Req 3.2/3.3),
    # and in that case the output is not marked complete.
    item_count_discs = [d for d in summary.discrepancies if d.kind == "ITEM_COUNT"]
    actual_count = len(doc.line_items)
    expects_item_count_disc = (
        doc.declared_item_count is not None
        and doc.declared_item_count != actual_count
    )
    if expects_item_count_disc:
        assert len(item_count_discs) == 1
        disc = item_count_discs[0]
        assert disc.expected == doc.declared_item_count
        assert disc.actual == actual_count
        assert result.output_complete is False
    else:
        assert item_count_discs == []
        assert result.output_complete is True
