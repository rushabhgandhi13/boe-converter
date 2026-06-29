"""Upload Validation: ordered size -> readable PDF -> BOE-recognition checks.

Implements task 9.1. ``UploadValidator.validate`` runs the three rejection
checks from the design's Error Handling table in order and returns the first
failure, or an OK outcome carrying an opened PDF handle.

Design references:
- "Upload Validation" component (ordered checks: size -> readable PDF -> BOE).
- "Rejections (no output produced)" table for the exact user messages.
- BOE recognition (Req 1.6) is a positive heuristic: the document must contain
  every BOE marker token; absence of any marker => NOT_A_BOE.

Requirements: 1.2, 1.5, 1.6
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Any

import pdfplumber

# Error codes for the rejection paths (surfaced as ``error_code`` by the API).
ERROR_FILE_TOO_LARGE = "FILE_TOO_LARGE"
ERROR_INVALID_PDF = "INVALID_PDF"
ERROR_NOT_A_BOE = "NOT_A_BOE"

# Exact user messages from the design's Rejections table (Req 1.2, 1.5, 1.6).
MESSAGE_FILE_TOO_LARGE = "File exceeds the 50 MB size limit."
MESSAGE_INVALID_PDF = "A valid PDF is required."
MESSAGE_NOT_A_BOE = "The document is not a recognized Bill of Entry."

# BOE marker tokens. A document is recognized when it matches *either* known
# layout's full marker set (Req 1.6): the new ICEGATE format or the legacy
# "Indian Customs EDI System" format. Absence of any marker in both sets =>
# NOT_A_BOE.
BOE_MARKERS = (
    "BILL OF ENTRY",
    "Port Code",
    "BE No",
    "PART - II",
    "PART - III",
)

# Legacy "Indian Customs EDI System - Imports V1.5R001" format markers.
BOE_MARKERS_OLD = (
    "BILL OF ENTRY",
    "Indian Customs EDI System",
    "BE No",
    "Item Details",
)

# All recognized marker sets; a match against any one set accepts the document.
BOE_MARKER_SETS = (BOE_MARKERS, BOE_MARKERS_OLD)


@dataclass
class ValidationOutcome:
    """Result of running the ordered upload checks.

    On success (``ok is True``) ``handle`` carries the opened ``pdfplumber.PDF``
    so downstream parsing does not need to re-open the document. On failure
    ``error_code`` and ``message`` describe the first check that failed and
    ``handle`` is ``None`` (no output is produced for a rejection).
    """

    ok: bool
    error_code: str | None = None
    message: str | None = None
    handle: Any | None = None  # pdfplumber.PDF when ok, else None

    @classmethod
    def success(cls, handle: Any) -> "ValidationOutcome":
        """Build an OK outcome carrying the opened PDF handle."""
        return cls(ok=True, error_code=None, message=None, handle=handle)

    @classmethod
    def failure(cls, error_code: str, message: str) -> "ValidationOutcome":
        """Build a rejection outcome with its error code and user message."""
        return cls(ok=False, error_code=error_code, message=message, handle=None)


class UploadValidator:
    """Validates an uploaded file before conversion begins."""

    MAX_BYTES = 50 * 1024 * 1024  # Req 1.2

    def validate(self, raw: bytes, filename: str) -> ValidationOutcome:
        """Run ordered checks and return the first failure, else OK.

        Order (design "Upload Validation"): size limit -> readable PDF ->
        recognizable BOE. The first applicable failure is returned so the most
        specific early message wins.
        """
        # 1) Size limit (Req 1.2).
        if len(raw) > self.MAX_BYTES:
            return ValidationOutcome.failure(ERROR_FILE_TOO_LARGE, MESSAGE_FILE_TOO_LARGE)

        # 2) Readable PDF (Req 1.5). Open and confirm at least one readable page.
        try:
            handle = pdfplumber.open(io.BytesIO(raw))
        except Exception:
            return ValidationOutcome.failure(ERROR_INVALID_PDF, MESSAGE_INVALID_PDF)

        try:
            pages = handle.pages
            if not pages:
                handle.close()
                return ValidationOutcome.failure(ERROR_INVALID_PDF, MESSAGE_INVALID_PDF)
        except Exception:
            # A file that opens but cannot be paginated is not a readable PDF.
            self._safe_close(handle)
            return ValidationOutcome.failure(ERROR_INVALID_PDF, MESSAGE_INVALID_PDF)

        # 3) Recognizable BOE (Req 1.6). All marker tokens must be present.
        try:
            if not self._is_recognized_boe(handle):
                handle.close()
                return ValidationOutcome.failure(ERROR_NOT_A_BOE, MESSAGE_NOT_A_BOE)
        except Exception:
            # If we cannot extract any text to scan, the document is not a
            # recognizable BOE (and is not silently accepted).
            self._safe_close(handle)
            return ValidationOutcome.failure(ERROR_NOT_A_BOE, MESSAGE_NOT_A_BOE)

        # All checks passed: return OK with the still-open handle.
        return ValidationOutcome.success(handle)

    # -- helpers ------------------------------------------------------------
    def _is_recognized_boe(self, handle: Any) -> bool:
        """True when the document matches any known BOE layout's marker set.

        Text is scanned page by page with whitespace collapsed and case
        normalized so tokens spanning line breaks or with irregular spacing are
        still matched. The markers seen across all pages are accumulated, and the
        document is recognized as soon as every marker of *any* one set (the new
        ICEGATE format or the legacy EDI format) has been observed.
        """
        marker_sets = [
            {self._normalize(m) for m in markers} for markers in BOE_MARKER_SETS
        ]
        for page in handle.pages:
            try:
                text = page.extract_text() or ""
            except Exception:
                continue
            norm = self._normalize(text)
            marker_sets = [{m for m in s if m not in norm} for s in marker_sets]
            if any(not s for s in marker_sets):
                return True
        return any(not s for s in marker_sets)

    @staticmethod
    def _normalize(text: str) -> str:
        """Collapse all whitespace to single spaces and upper-case the text."""
        return re.sub(r"\s+", " ", text).strip().upper()

    @staticmethod
    def _safe_close(handle: Any) -> None:
        """Close a PDF handle, ignoring any close-time errors."""
        try:
            handle.close()
        except Exception:
            pass
