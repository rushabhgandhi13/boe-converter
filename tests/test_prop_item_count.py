"""Property 2: Extracted item count is preserved and mismatches are reported.

*For any* extracted set of line items with a declared item count, the number of
records equals the number of distinct item serials; and *for any*
declared/extracted count pair that differs, a ``Discrepancy`` of kind
``ITEM_COUNT`` is produced carrying both the declared and extracted counts and
the output is not marked complete.

**Validates: Requirements 3.2, 3.3**

This exercises the ``ConversionOrchestrator`` cross-check end to end with
stubbed components so no real PDF parsing/Excel writing happens:

- a fake ``UploadValidator`` returning an OK ``ValidationOutcome`` (``ok=True``,
  ``handle=None``) so the pipeline proceeds past recognition;
- a fake ``PdfParser`` whose ``parse(handle, usd_rate=...)`` returns a generated
  ``ExtractedDocument`` with ``N`` line items (distinct serials) and a
  ``declared_item_count`` that is sometimes ``None``, sometimes ``== N``, and
  sometimes ``!= N``;
- the real ``ValueCalculator`` (pure, no I/O); and
- a fake ``ExcelGenerator`` returning ``b"xlsx"`` to avoid heavy ``openpyxl`` work.
"""

from __future__ import annotations

import hypothesis.strategies as st
from hypothesis import given

from boe_converter.models import (
    ExtractedDocument,
    HeaderBlock,
    LineItem,
    RawValue,
)
from boe_converter.orchestrator import ConversionOrchestrator
from boe_converter.validator import ValidationOutcome


# ---------------------------------------------------------------------------
# Stub components injected into the orchestrator
# ---------------------------------------------------------------------------
class _FakeValidator:
    """Always accepts: returns an OK outcome with no handle (Req 1.x bypassed)."""

    def validate(self, raw: bytes, filename: str) -> ValidationOutcome:
        return ValidationOutcome.success(None)


class _FakeParser:
    """Returns a pre-built ``ExtractedDocument`` regardless of the input."""

    def __init__(self, doc: ExtractedDocument) -> None:
        self._doc = doc

    def parse(self, handle, usd_rate: float) -> ExtractedDocument:
        return self._doc


class _FakeGenerator:
    """Returns fixed bytes, avoiding real workbook construction."""

    def generate(self, doc, flags) -> bytes:
        return b"xlsx"


# ---------------------------------------------------------------------------
# Generators
# ---------------------------------------------------------------------------
def _num(value: float) -> RawValue:
    """A located, parseable numeric field as it would arrive from the parser."""

    return RawValue(raw_text=str(value), parsed=value)


def _make_line_item(serial: int) -> LineItem:
    """A minimal, fully-valid line item (no review flags, no recompute issues)."""

    return LineItem(
        item_serial=serial,
        cth_hsn=RawValue(raw_text="1234", parsed="1234"),
        description=RawValue(raw_text="widget", parsed="widget"),
        unit_price_usd=_num(1.0),
        quantity=_num(1.0),
        unit=RawValue(raw_text="PCS", parsed="PCS"),
        assessable_value=_num(100.0),
        bcd_rate=_num(0.10),
        bcd_amount=_num(10.0),
        igst_rate=_num(0.18),
        total_duty=_num(0.0),
    )


def _make_header() -> HeaderBlock:
    """A header with an absent invoice amount (item-count check is isolated)."""

    missing = RawValue.missing()
    return HeaderBlock(
        company_name="Gemini Unicom LLP",
        party_name=missing,
        usd_rate=95.0,
        details=missing,
        invoice_no=missing,
        invoice_date=missing,
        be_no=missing,
        be_date=missing,
        bl_no=missing,
        bl_date=missing,
        invoice_amount=missing,
        invoice_currency=missing,
        package_count=missing,
        container_details=missing,
    )


@st.composite
def _scenario(draw):
    """Build (ExtractedDocument, extracted_count, declared, expect_mismatch).

    ``serials`` are unique so the record count equals the number of distinct
    item serials. ``declared`` independently lands on one of: absent (None),
    equal to the extracted count, or a differing non-negative integer.
    """

    serials = draw(
        st.lists(
            st.integers(min_value=1, max_value=500),
            min_size=0,
            max_size=10,
            unique=True,
        )
    )
    extracted = len(serials)

    mode = draw(st.sampled_from(("none", "match", "mismatch")))
    if mode == "none":
        declared = None
    elif mode == "match":
        declared = extracted
    else:
        declared = draw(
            st.integers(min_value=0, max_value=600).filter(lambda d: d != extracted)
        )

    doc = ExtractedDocument(
        header=_make_header(),
        line_items=[_make_line_item(s) for s in serials],
        declared_item_count=declared,
    )
    expect_mismatch = declared is not None and declared != extracted
    return doc, extracted, declared, expect_mismatch


def _orchestrator_for(doc: ExtractedDocument) -> ConversionOrchestrator:
    return ConversionOrchestrator(
        validator=_FakeValidator(),
        parser=_FakeParser(doc),
        generator=_FakeGenerator(),
    )


# ---------------------------------------------------------------------------
# Property
# ---------------------------------------------------------------------------
@given(scenario=_scenario())
def test_item_count_preserved_and_mismatches_reported(scenario):
    doc, extracted, declared, expect_mismatch = scenario

    result = _orchestrator_for(doc).convert(b"pdf-bytes", "boe.pdf", usd_rate=95.0)

    assert result.ok, "the conversion must succeed for a recognized BOE"
    summary = result.summary

    # (a) Count preserved: the summary reports exactly the parsed records, and
    #     they equal the number of distinct item serials (no drops, no merges).
    assert summary.line_items_extracted == extracted
    assert extracted == len({li.item_serial for li in doc.line_items})
    assert summary.declared_item_count == declared

    item_count_discs = [d for d in summary.discrepancies if d.kind == "ITEM_COUNT"]

    if expect_mismatch:
        # (b) A mismatch yields exactly one ITEM_COUNT discrepancy carrying both
        #     the declared and extracted counts, and the output is not complete.
        assert len(item_count_discs) == 1
        disc = item_count_discs[0]
        assert disc.expected == declared
        assert disc.actual == extracted
        assert result.output_complete is False
    else:
        # (c) Matching (or absent) declared count: no ITEM_COUNT discrepancy and
        #     the output is presented as complete.
        assert item_count_discs == []
        assert result.output_complete is True
