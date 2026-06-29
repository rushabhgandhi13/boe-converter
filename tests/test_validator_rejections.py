"""Unit tests for the UploadValidator rejection paths (task 9.2).

These tests exercise the three "no output produced" rejections from the
design's Error Handling table by driving :meth:`UploadValidator.validate`
with crafted inputs:

1. An oversized payload (``len(raw) > MAX_BYTES``) is rejected as
   ``FILE_TOO_LARGE`` before any PDF parsing is attempted (Req 1.2).
2. Bytes that are not a readable PDF are rejected as ``INVALID_PDF`` (Req 1.5).
3. A genuine, readable PDF that lacks the BOE marker tokens is rejected as
   ``NOT_A_BOE`` (Req 1.6).

In every rejection case the outcome must report ``ok is False``, the exact
user-facing message from the design, and carry no opened document handle (no
output is produced for a rejection).

Validates: Requirements 1.2, 1.5, 1.6
"""

from __future__ import annotations

import io

import pytest

from boe_converter.validator import (
    ERROR_FILE_TOO_LARGE,
    ERROR_INVALID_PDF,
    ERROR_NOT_A_BOE,
    MESSAGE_FILE_TOO_LARGE,
    MESSAGE_INVALID_PDF,
    MESSAGE_NOT_A_BOE,
    UploadValidator,
)

# reportlab is used to synthesise a real, readable non-BOE PDF in memory.
reportlab_canvas = pytest.importorskip("reportlab.pdfgen.canvas")


def _make_non_boe_pdf() -> bytes:
    """Build a minimal, valid one-page PDF that is NOT a Bill of Entry.

    The page carries ordinary prose containing none of the BOE marker tokens
    (``BILL OF ENTRY``, ``Port Code``, ``BE No``, ``PART - II``, ``PART - III``)
    so the document opens and paginates cleanly yet fails BOE recognition.
    """
    buffer = io.BytesIO()
    pdf = reportlab_canvas.Canvas(buffer)
    pdf.drawString(100, 750, "This is an ordinary invoice, not a customs document.")
    pdf.drawString(100, 730, "Thank you for your business. Total due: 42.00 dollars.")
    pdf.showPage()
    pdf.save()
    return buffer.getvalue()


def _assert_rejected(outcome, error_code: str, message: str) -> None:
    """Assert a rejection outcome: not ok, exact code/message, no handle."""
    assert outcome.ok is False
    assert outcome.error_code == error_code
    assert outcome.message == message
    assert outcome.handle is None


def test_oversized_file_rejected_with_size_message():
    """Req 1.2 - a payload larger than MAX_BYTES yields FILE_TOO_LARGE."""
    validator = UploadValidator()
    # One byte over the limit is enough to trip the size check; the content
    # itself need not be a PDF because the size check runs first.
    oversized = b"\x00" * (UploadValidator.MAX_BYTES + 1)

    outcome = validator.validate(oversized, "huge.pdf")

    _assert_rejected(outcome, ERROR_FILE_TOO_LARGE, MESSAGE_FILE_TOO_LARGE)
    assert outcome.message == "File exceeds the 50 MB size limit."


def test_corrupt_non_pdf_bytes_rejected_as_invalid_pdf():
    """Req 1.5 - bytes that are not a readable PDF yield INVALID_PDF."""
    validator = UploadValidator()
    corrupt = b"not a pdf"

    outcome = validator.validate(corrupt, "broken.pdf")

    _assert_rejected(outcome, ERROR_INVALID_PDF, MESSAGE_INVALID_PDF)
    assert outcome.message == "A valid PDF is required."


def test_valid_non_boe_pdf_rejected_as_not_a_boe():
    """Req 1.6 - a readable PDF lacking BOE markers yields NOT_A_BOE."""
    validator = UploadValidator()
    non_boe_pdf = _make_non_boe_pdf()

    outcome = validator.validate(non_boe_pdf, "invoice.pdf")

    _assert_rejected(outcome, ERROR_NOT_A_BOE, MESSAGE_NOT_A_BOE)
    assert outcome.message == "The document is not a recognized Bill of Entry."


def test_empty_bytes_rejected_as_invalid_pdf():
    """Req 1.5 - an empty payload is within the size limit but not a readable
    PDF, so it is rejected as INVALID_PDF rather than accepted."""
    validator = UploadValidator()

    outcome = validator.validate(b"", "empty.pdf")

    _assert_rejected(outcome, ERROR_INVALID_PDF, MESSAGE_INVALID_PDF)
