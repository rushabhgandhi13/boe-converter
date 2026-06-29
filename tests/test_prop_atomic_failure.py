"""Property-based test for atomic output on post-recognition failure.

Property 17: Failure after BOE recognition yields no downloadable output.

**Validates: Requirements 1.7**

Strategy: drive the ``ConversionOrchestrator`` with injected fakes. The fake
``UploadValidator`` always returns ``ok=True`` (the document is recognized as a
BOE), so every failure exercised here is *post-recognition*. Hypothesis then
chooses which post-recognition stage fails -- ``parser.parse``,
``calculator.compute``, or ``generator.generate`` -- and varies the exception
type and message. For every failing stage the property asserts:

- ``convert(...)`` returns ``ok is False`` with ``error_code`` equal to
  ``ERROR_CONVERSION_FAILED`` and ``download_token is None``; and
- no workbook is downloadable: the internal ``_downloads`` store stays empty and
  ``get_download(token)`` is ``None`` for any token (no partial output leaks).

A complementary success case (no failure injected: fake validator + fake parser
feeding a *real* ``ValueCalculator`` and a fake generator returning ``b"xlsx"``)
confirms the token store only populates on a fully successful build, and that
``get_download`` then returns exactly those bytes.
"""

from __future__ import annotations

import secrets

from hypothesis import given
from hypothesis import strategies as st

from boe_converter.calculator import ValueCalculator
from boe_converter.models import (
    ComputedDocument,
    ExtractedDocument,
    HeaderBlock,
    RawValue,
    ReviewFlagSet,
    Totals,
)
from boe_converter.orchestrator import (
    ERROR_CONVERSION_FAILED,
    ConversionOrchestrator,
)
from boe_converter.validator import ValidationOutcome


# ---------------------------------------------------------------------------
# Minimal domain fixtures used by the fakes
# ---------------------------------------------------------------------------


def _header() -> HeaderBlock:
    """A minimal, fully-populated header for a recognized BOE (no line items)."""

    return HeaderBlock(
        company_name="Gemini Unicom LLP",
        party_name=RawValue.missing(),
        usd_rate=95.0,
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


def _extracted() -> ExtractedDocument:
    """An empty-but-valid extraction result (BOE recognized, no line items)."""

    return ExtractedDocument(header=_header(), line_items=[], declared_item_count=0)


def _computed() -> ComputedDocument:
    """A trivial computed document for the fake generator's input."""

    return ComputedDocument(header=_header(), lines=[], totals=Totals())


class _FakeHandle:
    """Stands in for the opened ``pdfplumber.PDF`` handle; tracks close()."""

    def __init__(self) -> None:
        self.closed = False

    def close(self) -> None:
        self.closed = True


# ---------------------------------------------------------------------------
# Injected fakes
# ---------------------------------------------------------------------------


class _FakeValidator:
    """Always recognizes the upload as a BOE (ok=True)."""

    def validate(self, raw: bytes, filename: str) -> ValidationOutcome:
        return ValidationOutcome.success(_FakeHandle())


class _FakeParser:
    """Returns an extraction result, or raises if configured to fail."""

    def __init__(self, exc: Exception | None = None) -> None:
        self._exc = exc

    def parse(self, handle, usd_rate: float) -> ExtractedDocument:
        if self._exc is not None:
            raise self._exc
        return _extracted()


class _FakeCalculator:
    """Returns a computed document, or raises if configured to fail."""

    def __init__(self, exc: Exception | None = None) -> None:
        self._exc = exc

    def compute(self, extracted: ExtractedDocument, usd_rate: float) -> ComputedDocument:
        if self._exc is not None:
            raise self._exc
        return _computed()


class _FakeGenerator:
    """Returns workbook bytes, or raises if configured to fail."""

    def __init__(self, exc: Exception | None = None, data: bytes = b"xlsx") -> None:
        self._exc = exc
        self._data = data

    def generate(self, computed: ComputedDocument, flags: ReviewFlagSet) -> bytes:
        if self._exc is not None:
            raise self._exc
        return self._data


# ---------------------------------------------------------------------------
# Strategies
# ---------------------------------------------------------------------------

# Which post-recognition stage fails.
_STAGE = st.sampled_from(("parse", "compute", "generate"))

# Vary the exception class and message raised at the failing stage.
_EXC_TYPE = st.sampled_from(
    (ValueError, RuntimeError, KeyError, TypeError, ZeroDivisionError, Exception)
)
_MESSAGE = st.text(max_size=40)


def _make_exception(exc_type, message: str) -> Exception:
    """Build an exception instance of ``exc_type`` carrying ``message``."""

    return exc_type(message)


# ---------------------------------------------------------------------------
# Property 17 -- post-recognition failure yields no downloadable output
# ---------------------------------------------------------------------------


@given(
    stage=_STAGE,
    exc_type=_EXC_TYPE,
    message=_MESSAGE,
    usd_rate=st.floats(
        min_value=0.0, max_value=1e6, allow_nan=False, allow_infinity=False
    ),
)
def test_post_recognition_failure_yields_no_download(
    stage, exc_type, message, usd_rate
):
    """Any exception after BOE recognition => failure, no token, empty store."""

    exc = _make_exception(exc_type, message)

    parser = _FakeParser(exc if stage == "parse" else None)
    calculator = _FakeCalculator(exc if stage == "compute" else None)
    generator = _FakeGenerator(exc if stage == "generate" else None)

    orchestrator = ConversionOrchestrator(
        validator=_FakeValidator(),
        parser=parser,
        calculator=calculator,
        generator=generator,
    )

    result = orchestrator.convert(b"%PDF-fake-bytes", "boe.pdf", usd_rate)

    # The conversion is reported as failed with the post-recognition error code.
    assert result.ok is False
    assert result.error_code == ERROR_CONVERSION_FAILED
    # No token references any (partial) output.
    assert result.download_token is None
    assert result.summary is None

    # Atomic output: nothing was stored, so no workbook is downloadable.
    assert orchestrator._downloads == {}
    # Any token -- including a freshly minted random one -- yields nothing.
    assert orchestrator.get_download(secrets.token_urlsafe(32)) is None


# ---------------------------------------------------------------------------
# Success case -- the token store only populates on a fully successful build
# ---------------------------------------------------------------------------


@given(
    usd_rate=st.floats(
        min_value=0.0, max_value=1e6, allow_nan=False, allow_infinity=False
    ),
)
def test_successful_build_issues_token_and_serves_bytes(usd_rate):
    """No failure injected: a token is issued and get_download returns the bytes."""

    orchestrator = ConversionOrchestrator(
        validator=_FakeValidator(),
        parser=_FakeParser(),
        calculator=ValueCalculator(),  # real calculator on an empty document
        generator=_FakeGenerator(data=b"xlsx"),
    )

    result = orchestrator.convert(b"%PDF-fake-bytes", "boe.pdf", usd_rate)

    assert result.ok is True
    assert result.error_code is None
    assert result.download_token is not None
    # The token store populated exactly once, on success.
    assert len(orchestrator._downloads) == 1
    # The issued token serves precisely the generated workbook bytes.
    assert orchestrator.get_download(result.download_token) == b"xlsx"
