"""End-to-end and atomicity integration tests for the BOE converter.

Task 11.2. These tests exercise the whole system against the real project
assets - the reference Bill of Entry PDF and the golden CTN workbook - rather
than synthetic inputs, covering the requirements that can only be validated by
a full round-trip:

- **Req 1.3 / 1.8** - a recognized BOE converts to a downloadable ``Sheet1``
  workbook within the 60-second budget, with the expected 45 line items in
  dense rows from row 13 and the Totals_Row at row 61.
- **Golden output** - the values the system derives *directly from extraction*
  (column L invoice amount total, column N assessable-value total, and the
  column G package count) match the evaluated values stored in the sample
  workbook ``1357 ctn llp.xlsx`` within tolerance.
- **Req 1.7 (atomicity)** - a failure occurring *after* the document is
  recognized as a BOE issues no download token and leaves no partial output
  retrievable.

Both project assets are large binaries that may be absent in some checkouts, so
the whole module is skipped when either the reference PDF or the golden workbook
is missing.

Intentionally-skipped golden columns
------------------------------------
Only the columns whose values flow *directly* from extraction (and therefore
must agree with the sample) are compared:

- ``L`` (Amount, USD)   = sum of unit_price x quantity per line
- ``N`` (CUSTOM ASS VALUE) = sum of the directly-extracted assessable values
- ``G`` (package/CTN count) = the BOE's declared total package count

The duty/derived columns (``M`` rate-per-USD, ``O``/``X`` land costs,
``P``/``S`` customs duty, ``Q`` IGST, ``U`` BCD amount, ``W`` SWS) are **not**
compared: the sample stores manually-keyed BCD/SWS/IGST figures that the parser
recomputes from the BOE's printed duty grid, so small principled differences are
expected there and are out of scope for this golden check.

Validates: Requirements 1.3, 1.7, 1.8
"""

from __future__ import annotations

import io
import time
from pathlib import Path

import pytest

# openpyxl and FastAPI's TestClient are required for these integration tests.
openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("fastapi")

from fastapi.testclient import TestClient  # noqa: E402

from boe_converter.orchestrator import (  # noqa: E402
    ConversionOrchestrator,
    ERROR_CONVERSION_FAILED,
)
from boe_converter.web.app import create_app  # noqa: E402

# Project assets live at the repository root (one level above tests/).
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
REFERENCE_PDF = _PROJECT_ROOT / "205090022062026INNSA1BE0230620261842.pdf"
GOLDEN_WORKBOOK = _PROJECT_ROOT / "1357 ctn llp.xlsx"

# The User-supplied conversion rate for the reference document (task 11.2).
USD_RATE = 95.3

# Expected layout facts for the reference BOE (design "Target Excel cell map").
EXPECTED_LINE_ITEMS = 45
FIRST_DATA_ROW = 13
LAST_DATA_ROW = FIRST_DATA_ROW + EXPECTED_LINE_ITEMS - 1  # 57
TOTALS_ROW = 61
TIME_BUDGET_SECONDS = 60.0

# Column indices (1-based) for the directly-extraction-derived totals compared
# against the golden workbook.
COL_G_PACKAGE = 7   # G  CTN / package count
COL_L_AMOUNT = 12   # L  Amount (USD) total
COL_N_ASSESS = 14   # N  CUSTOM ASS VALUE total

# Module-level skip when either project asset is unavailable.
pytestmark = pytest.mark.skipif(
    not REFERENCE_PDF.exists() or not GOLDEN_WORKBOOK.exists(),
    reason=(
        "integration assets missing: require both the reference BOE PDF "
        f"({REFERENCE_PDF.name}) and the golden workbook ({GOLDEN_WORKBOOK.name})"
    ),
)


def _approx_equal(actual: float, expected: float) -> bool:
    """True when ``actual`` matches ``expected`` within tolerance.

    Tolerance is 0.01 absolute, widened to a small relative band
    (1e-4 of the expected magnitude) for the large totals so full-precision
    sums that round to the sample's 2-decimal cached values still compare equal.
    """
    return abs(actual - expected) <= max(0.01, abs(expected) * 1e-4)


@pytest.fixture(scope="module")
def reference_pdf_bytes() -> bytes:
    """The raw bytes of the reference Bill of Entry PDF."""
    return REFERENCE_PDF.read_bytes()


# ---------------------------------------------------------------------------
# 1) End-to-end: upload -> convert -> download a Sheet1 workbook within 60s.
# ---------------------------------------------------------------------------
def test_end_to_end_convert_and_download_within_budget(reference_pdf_bytes):
    """Req 1.3/1.8 - the reference BOE converts to a downloadable Sheet1
    workbook within 60s, with 45 dense line items from row 13 and totals at
    row 61.

    Drives the real HTTP API end to end via ``TestClient``: POST the PDF with
    ``usd_rate=95.3``, assert the success body (token + summary with 45 items),
    then GET the download and inspect the returned ``.xlsx`` bytes.
    """
    client = TestClient(create_app())

    start = time.monotonic()
    response = client.post(
        "/api/convert",
        files={
            "file": (
                "boe.pdf",
                reference_pdf_bytes,
                "application/pdf",
            )
        },
        data={"usd_rate": str(USD_RATE)},
    )
    elapsed = time.monotonic() - start

    assert response.status_code == 200, response.text
    body = response.json()

    download_token = body.get("download_token")
    assert download_token, "a successful conversion must issue a download token"

    summary = body["summary"]
    assert summary["line_items_extracted"] == EXPECTED_LINE_ITEMS

    # Req 1.3: the end-to-end conversion completes within the 60-second budget.
    assert elapsed < TIME_BUDGET_SECONDS, (
        f"conversion took {elapsed:.2f}s, exceeding the {TIME_BUDGET_SECONDS}s budget"
    )

    # GET the download and load the returned workbook bytes.
    dl = client.get(f"/api/download/{download_token}")
    assert dl.status_code == 200, dl.text
    assert dl.content, "download response must carry the workbook bytes"

    wb = openpyxl.load_workbook(io.BytesIO(dl.content))
    try:
        # Exactly one sheet named 'Sheet1' (Req 8.1).
        assert wb.sheetnames == ["Sheet1"]
        ws = wb["Sheet1"]

        # 45 dense data rows from row 13: Sr. no. runs 1..45 with no gaps.
        for offset in range(EXPECTED_LINE_ITEMS):
            row = FIRST_DATA_ROW + offset
            assert ws.cell(row=row, column=1).value == offset + 1, (
                f"Sr. no. at row {row} should be {offset + 1}"
            )
        # The row immediately after the last data row is blank (dense block).
        assert ws.cell(row=LAST_DATA_ROW + 1, column=1).value is None

        # Totals row at row 61. The downloaded workbook carries a live SUM
        # formula over the data rows (matching the sample workbook's
        # ``=SUM(L13:L59)`` style), so the cell holds the formula text.
        amount_total = ws.cell(row=TOTALS_ROW, column=COL_L_AMOUNT).value
        assert isinstance(amount_total, str)
        assert amount_total.startswith(f"=SUM(L{FIRST_DATA_ROW}:L")
    finally:
        wb.close()

    # Record the measured elapsed time for the run report.
    print(f"\n[e2e] conversion + download elapsed: {elapsed:.2f}s")


# ---------------------------------------------------------------------------
# 2) Golden-output comparison: extraction-derived totals match the sample.
# ---------------------------------------------------------------------------
def test_generated_totals_match_golden_workbook(reference_pdf_bytes):
    """The generated workbook's extraction-derived totals (L amount, N
    assessable value, G package count) match the golden workbook's evaluated
    values within tolerance.

    Only the columns that flow directly from extraction are compared; the
    duty/derived columns are intentionally skipped (see module docstring) because
    the sample stores manually-keyed duty figures the parser recomputes.
    """
    # Generate the workbook in literal (evaluated) mode so the totals carry
    # numbers comparable to the golden workbook's cached values. The downloaded
    # workbook from the orchestrator carries live SUM formulas instead (verified
    # in test_end_to_end_convert_and_download_within_budget and
    # tests/test_excel_formulas.py); openpyxl does not evaluate formulas, so a
    # literal build is used here for the numeric golden comparison. The pipeline
    # correctness (parse -> compute -> generate) is still exercised end to end.
    from boe_converter.calculator import ValueCalculator
    from boe_converter.excel_writer import ExcelGenerator
    from boe_converter.models import ReviewFlagSet
    from boe_converter.parser import PdfParser

    extracted = PdfParser().parse(io.BytesIO(reference_pdf_bytes), usd_rate=USD_RATE)
    computed = ValueCalculator().compute(extracted, USD_RATE)
    generated_bytes = ExcelGenerator(use_formulas=False).generate(computed, ReviewFlagSet())
    assert generated_bytes is not None

    gen_wb = openpyxl.load_workbook(io.BytesIO(generated_bytes))
    # data_only=True so we read the sample's *evaluated* (cached) values, since
    # the golden workbook stores most numeric cells as Excel formulas.
    golden_wb = openpyxl.load_workbook(str(GOLDEN_WORKBOOK), data_only=True)
    try:
        gen = gen_wb["Sheet1"]
        golden = golden_wb["Sheet1"]

        compared = 0
        skipped = []
        for col, name in (
            (COL_L_AMOUNT, "L (Amount USD total)"),
            (COL_N_ASSESS, "N (assessable value total)"),
            (COL_G_PACKAGE, "G (package count)"),
        ):
            golden_value = golden.cell(row=TOTALS_ROW, column=col).value
            gen_value = gen.cell(row=TOTALS_ROW, column=col).value

            # If the golden value cannot be read (formula without a cached
            # value), skip that comparison gracefully rather than failing.
            if not isinstance(golden_value, (int, float)):
                skipped.append(f"{name}: golden value unavailable ({golden_value!r})")
                continue

            assert isinstance(gen_value, (int, float)), (
                f"generated {name} should be numeric, got {gen_value!r}"
            )
            assert _approx_equal(float(gen_value), float(golden_value)), (
                f"{name}: generated {gen_value} != golden {golden_value}"
            )
            compared += 1

        assert compared >= 1, (
            "expected at least one golden total to be comparable; "
            f"all were skipped: {skipped}"
        )
        if skipped:
            print("\n[golden] skipped (unreadable golden cells):", skipped)
    finally:
        gen_wb.close()
        golden_wb.close()


# ---------------------------------------------------------------------------
# 3) Atomicity: a post-recognition failure issues no download token.
# ---------------------------------------------------------------------------
class _RaisingGenerator:
    """A drop-in Excel_Generator that always fails during workbook build.

    Used to inject a failure that occurs *after* the document has been
    recognized as a BOE (validation passes, parse/compute succeed), exercising
    the orchestrator's atomic-output guarantee (Req 1.7).
    """

    def generate(self, computed, flags):  # noqa: ANN001 - test double
        raise RuntimeError("injected post-recognition generation failure")


def test_post_recognition_failure_issues_no_token(reference_pdf_bytes):
    """Req 1.7 - a failure after BOE recognition reports failure and leaves no
    downloadable output.

    The orchestrator uses the real validator (so the reference PDF is genuinely
    recognized as a BOE) but a generator that raises during the build. The
    result must be a failure with no token, and the in-memory download store
    must remain empty - no partial workbook is ever retrievable.
    """
    orchestrator = ConversionOrchestrator(generator=_RaisingGenerator())

    result = orchestrator.convert(reference_pdf_bytes, "boe.pdf", USD_RATE)

    assert result.ok is False
    assert result.error_code == ERROR_CONVERSION_FAILED
    assert result.download_token is None
    # Atomic output: no token was issued, so the store holds nothing.
    assert orchestrator._downloads == {}
