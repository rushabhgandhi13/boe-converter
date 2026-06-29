"""Orientation-robustness regression test (task 4.3).

This test pins the orientation-aware extraction primitive against the *real*
reference BOE PDF. The BOE renders rotated/vertical section labels in the page
margins (e.g. ``SLIATED``, ``YTUD``, ``SEITUD`` -- the reversed glyph runs of
``DETAILS``/``DUTY``/``DUTIES``). Naive full-page ``extract_text`` interleaves
this rotated text with the real horizontal data, producing garbage rows.

``PdfParser._upright_words`` keeps only characters whose text rotation is 0
degrees, so those rotated margin labels must never appear in the reconstructed
upright text. We assert both halves of that contract:

1. A naive ``page.extract_text()`` on page 0 DOES contain a rotated label token
   (proving the hazard is real for this document, i.e. the test has teeth).
2. The joined text of ``PdfParser()._upright_words(page)`` does NOT contain any
   of those rotated tokens (proving the filter defeats the hazard).

The reference PDF is large binary data that may not be present in every
checkout, so the whole module is skipped when it is absent.

_Requirements: 3.6_
"""

from __future__ import annotations

from pathlib import Path

import pytest

pdfplumber = pytest.importorskip("pdfplumber")

from boe_converter.parser import PdfParser

# The reference BOE PDF lives at the project root (two levels up from this file:
# tests/ -> project root).
_REFERENCE_PDF = (
    Path(__file__).resolve().parents[1]
    / "205090022062026INNSA1BE0230620261842.pdf"
)

# Rotated margin labels observed on the BOE pages. These are reversed glyph runs
# of DETAILS / DUTY / DUTIES rendered vertically in the margin; they must be
# discarded by orientation filtering.
_ROTATED_TOKENS = ("SLIATED", "YTUD", "SEITUD")

pytestmark = pytest.mark.skipif(
    not _REFERENCE_PDF.exists(),
    reason=f"reference BOE PDF not available at {_REFERENCE_PDF}",
)


def _naive_text_page0() -> str:
    with pdfplumber.open(str(_REFERENCE_PDF)) as pdf:
        return pdf.pages[0].extract_text() or ""


def _upright_text_page0() -> str:
    parser = PdfParser()
    with pdfplumber.open(str(_REFERENCE_PDF)) as pdf:
        words = parser._upright_words(pdf.pages[0])
    return " ".join(w.text for w in words)


def test_naive_extraction_contains_rotated_label():
    """Naive extract_text is contaminated by at least one rotated margin label.

    This guards the regression test itself: if the source PDF or pdfplumber ever
    stops surfacing the rotated labels in naive extraction, the orientation
    filter would no longer be exercised and this test would give false comfort.
    """
    naive = _naive_text_page0()
    contaminating = [tok for tok in _ROTATED_TOKENS if tok in naive]
    assert contaminating, (
        "expected naive extract_text() on page 0 to contain a rotated margin "
        f"label from {_ROTATED_TOKENS!r}; got none, so the orientation filter "
        "is not being meaningfully exercised by this document"
    )


def test_upright_words_excludes_rotated_labels():
    """_upright_words drops every rotated margin label (the core contract).

    Rotated text must never contaminate the reconstructed upright rows, so none
    of the rotated tokens may appear in the joined upright-word text.
    """
    upright = _upright_text_page0()
    leaked = [tok for tok in _ROTATED_TOKENS if tok in upright]
    assert not leaked, (
        f"rotated margin label(s) {leaked!r} leaked into _upright_words output; "
        "rotated text must be filtered out so it cannot contaminate extracted rows"
    )


def test_upright_words_still_returns_real_horizontal_content():
    """Sanity: filtering does not throw away the real upright BOE content.

    A document with rotated labels stripped should still yield substantial
    horizontal text (the BOE marker is present on every page), confirming we
    excluded only the rotated tokens rather than everything.
    """
    upright = _upright_text_page0().upper()
    assert "BILL OF ENTRY" in upright
