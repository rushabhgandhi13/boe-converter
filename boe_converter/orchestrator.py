"""Conversion Orchestrator: sequences validate -> parse -> compute -> generate.

Implements task 9.3. The orchestrator is the single coordinator that:

- runs the ordered upload checks (``UploadValidator``) and, on rejection,
  surfaces the first failure's ``error_code``/``message`` with **no output**
  (Req 1.2/1.5/1.6);
- sequences ``parse -> compute -> generate`` for a recognized BOE, aggregating
  every ``ReviewFlag`` raised across extraction and computation (Req 9.1);
- runs the cross-checks that compare the *declared* values printed in the BOE
  against the values the system *extracted/computed*:
    * ``Discrepancy(ITEM_COUNT)`` when the extracted line count differs from the
      BOE's declared count -- the output is then marked **not complete**
      (Req 3.2/3.3);
    * ``Discrepancy(INVOICE_TOTAL)`` when the computed sum of per-line USD
      amounts differs from the declared header invoice amount by more than
      0.01 USD -- the workbook is **retained** (Req 7.6); and a
      "could not be verified" message when the declared total is
      absent/unreadable (Req 7.7);
    * ``Discrepancy(RECOMPUTE)`` when a value extracted from the BOE (e.g. a
      line's ``total_duty``) differs from the value recomputed from its related
      fields by more than 0.01 (Req 9.4);
- builds the ``ConversionSummary`` rendered by the Web_Interface (Req 9.1);
- enforces **atomic output**: the generated ``.xlsx`` bytes are built entirely
  in memory and a ``download_token`` is issued *only* after a fully successful
  build, so a failure after BOE recognition never leaves a partial workbook
  downloadable (Req 1.7);
- enforces the **60-second budget**: an overrun (or any exception thrown after
  the document was recognized as a BOE) is treated as a post-recognition
  failure with no token issued (Req 1.3/1.7).

Design references: design.md "Conversion Orchestrator", "Error Handling"
(rejections vs anomalies), and the Web_Interface contract
(``{ download_token, summary }`` on success; mapped error bodies otherwise).
"""

from __future__ import annotations

import logging
import math
import secrets
import time
from dataclasses import dataclass, field, replace

from boe_converter.calculator import ValueCalculator
from boe_converter.excel_writer import ExcelGenerator
from boe_converter.invoice_parser import InvoicePackingListParser
from boe_converter.models import (
    GRAND_TOTAL_TOLERANCE_INR,
    ComputedDocument,
    ConversionSummary,
    Discrepancy,
    ExtractedDocument,
    GrandTotalCheck,
    GrandTotalVerification,
    LineItem,
    RawValue,
    ReviewFlag,
    ReviewFlagSet,
)
from boe_converter.parser import PdfParser
from boe_converter.validator import UploadValidator

logger = logging.getLogger(__name__)

# Error code/message for a post-recognition failure (Req 1.7). Rejection
# error codes/messages come from the validator (Req 1.2/1.5/1.6).
ERROR_CONVERSION_FAILED = "CONVERSION_FAILED"
MESSAGE_CONVERSION_FAILED = "The conversion could not be completed."

# Tolerance for the invoice-total (Req 7.6) and recompute (Req 9.4) checks.
TOLERANCE = 0.01

# Message wording for the declared-total-absent case (Req 7.7).
MESSAGE_TOTAL_UNVERIFIABLE = (
    "The computed invoice total could not be verified against the BOE."
)


@dataclass(frozen=True)
class ConversionResult:
    """Outcome of a single conversion request.

    Exactly one of two shapes is populated:

    - **Success** (``ok is True``): ``download_token`` references the generated
      ``.xlsx`` bytes (retrievable via :meth:`ConversionOrchestrator.get_download`)
      and ``summary`` carries the items/total/flags/discrepancies for the UI.
      ``output_complete`` is ``False`` when an item-count discrepancy means the
      output must not be presented as complete (Req 3.3).
    - **Failure** (``ok is False``): ``error_code``/``message`` describe a
      rejection (size/PDF/BOE-recognition, Req 1.2/1.5/1.6) or a post-recognition
      failure (``CONVERSION_FAILED``, Req 1.7); no token is ever issued.
    """

    ok: bool
    error_code: str | None = None
    message: str | None = None
    download_token: str | None = None
    summary: ConversionSummary | None = None
    output_complete: bool = True

    @classmethod
    def rejected(cls, error_code: str, message: str) -> "ConversionResult":
        """Build a rejection/failure result carrying no downloadable output."""
        return cls(ok=False, error_code=error_code, message=message)

    @classmethod
    def succeeded(
        cls,
        download_token: str,
        summary: ConversionSummary,
        *,
        output_complete: bool,
    ) -> "ConversionResult":
        """Build a success result referencing the issued download token."""
        return cls(
            ok=True,
            download_token=download_token,
            summary=summary,
            output_complete=output_complete,
        )


def _as_number(rv: RawValue | None) -> float | None:
    """Return the numeric interpretation of a ``RawValue`` (else ``None``).

    A value that is absent, unparseable, or whose ``parsed`` interpretation is
    not a real number yields ``None`` so the cross-checks treat it as
    "not available" rather than guessing.
    """
    if rv is None or rv.is_missing or rv.is_unparseable:
        return None
    parsed = rv.parsed
    if isinstance(parsed, bool):
        return None
    if isinstance(parsed, (int, float)):
        return float(parsed)
    return None


def _detail_number(detail: dict, key: str) -> float | None:
    """Numeric invoice value for ``key`` (else ``None``).

    Guards against hand-built detail mappings in tests carrying plain numbers:
    only a ``RawValue`` is interpreted, anything else counts as absent.
    """
    value = detail.get(key)
    return _as_number(value) if isinstance(value, RawValue) else None


class ConversionOrchestrator:
    """Coordinates the conversion pipeline end to end."""

    # Req 1.3: a recognized BOE must convert within 60 seconds; an overrun is a
    # post-recognition failure (Req 1.7) with no partial output.
    TIME_BUDGET_SECONDS = 60.0

    def __init__(
        self,
        validator: UploadValidator | None = None,
        parser: PdfParser | None = None,
        calculator: ValueCalculator | None = None,
        generator: ExcelGenerator | None = None,
        time_budget_seconds: float | None = None,
        invoice_parser: InvoicePackingListParser | None = None,
    ) -> None:
        self.validator = validator or UploadValidator()
        self.parser = parser or PdfParser()
        self.calculator = calculator or ValueCalculator()
        self.generator = generator or ExcelGenerator()
        self.invoice_parser = invoice_parser or InvoicePackingListParser()
        # The post-recognition time budget (Req 1.3 default 60s). It is made
        # configurable so a slower deployment (e.g. a shared-CPU Streamlit host
        # converting a large multi-hundred-page BOE) can allow more time rather
        # than failing a conversion that would otherwise succeed.
        self.time_budget_seconds = (
            self.TIME_BUDGET_SECONDS if time_budget_seconds is None else time_budget_seconds
        )
        # In-memory token -> workbook bytes store. A token is only added here
        # after a fully successful build, guaranteeing atomic output (Req 1.7).
        self._downloads: dict[str, bytes] = {}
        # Parallel token -> ComputedDocument store so the (separate) Tally-JSON
        # feature can build its export from the exact same in-memory computed
        # result the workbook was built from, without re-parsing the BOE.
        self._computed: dict[str, ComputedDocument] = {}

    def convert(
        self,
        raw: bytes,
        filename: str,
        usd_rate: float,
        invoice_raw: bytes | None = None,
    ) -> ConversionResult:
        """Run the full pipeline and return a :class:`ConversionResult`.

        Order of operations (design "Request flow"):

        1. Validate (size -> readable PDF -> recognizable BOE). A rejection
           returns immediately with the validator's message and no output.
        2. Parse -> compute -> generate the workbook entirely in memory. Any
           exception here is a post-recognition failure (Req 1.7).
        3. Enforce the 60-second budget; an overrun is a post-recognition
           failure with no token issued (Req 1.3/1.7).
        4. Run the declared-vs-extracted cross-checks (item count, invoice
           total, per-line recompute) and build the ``ConversionSummary``.
        5. Issue a ``download_token`` mapping to the workbook bytes and return
           success (Req 9.1).
        """
        start = time.monotonic()

        # 1) Ordered upload validation. First failure wins; no output produced.
        outcome = self.validator.validate(raw, filename)
        if not outcome.ok:
            return ConversionResult.rejected(
                outcome.error_code or ERROR_CONVERSION_FAILED,
                outcome.message or MESSAGE_CONVERSION_FAILED,
            )

        # 2) Parse -> compute -> generate. Everything past here is
        #    "post-recognition", so any failure is CONVERSION_FAILED (Req 1.7).
        handle = outcome.handle
        try:
            extracted = self.parser.parse(handle, usd_rate=usd_rate)
            # Optional: attach per-line carton counts from the supplier invoice.
            # Non-fatal - a missing/unreadable invoice simply leaves CTN blank.
            if invoice_raw:
                extracted = self._attach_cartons(extracted, invoice_raw)
            computed = self.calculator.compute(extracted, usd_rate)
            # Grand-total verification (Req 11) is computed *before* the
            # workbook is built so its MISMATCH columns can be red-filled by
            # generate(); the same GrandTotalVerification is reused for the
            # summary discrepancies below (computed once).
            verification = self._verify_grand_totals(extracted, computed)
            workbook_bytes = self.generator.generate(
                computed,
                ReviewFlagSet(computed.flags),
                grand_total_check=verification,
            )
        except Exception:
            # The document was recognized as a BOE, so this is a
            # post-recognition failure (Req 1.7): report a generic failure with
            # no token. Log the real cause so it is diagnosable from the server
            # / Streamlit logs (the user-facing message stays generic).
            logger.exception("Conversion failed after BOE recognition (file=%s)", filename)
            return ConversionResult.rejected(
                ERROR_CONVERSION_FAILED, MESSAGE_CONVERSION_FAILED
            )
        finally:
            self._safe_close(handle)

        # 3) Time budget (Req 1.3, configurable). Overrun => post-recognition
        #    failure (Req 1.7); the in-memory workbook is discarded, no token.
        elapsed = time.monotonic() - start
        if elapsed > self.time_budget_seconds:
            logger.warning(
                "Conversion exceeded the %.0fs budget (took %.1fs, file=%s); "
                "consider raising time_budget_seconds for this deployment.",
                self.time_budget_seconds, elapsed, filename,
            )
            return ConversionResult.rejected(
                ERROR_CONVERSION_FAILED, MESSAGE_CONVERSION_FAILED
            )

        # 4) Cross-checks + summary (workbook is retained regardless; these are
        #    anomalies surfaced to the User, never silent drops).
        discrepancies, output_complete = self._cross_check(
            extracted, computed, verification
        )
        summary = self._build_summary(extracted, computed, discrepancies)

        # 5) Atomic output: only now -- after a fully successful build -- issue
        #    the token mapping to the generated bytes (Req 1.7).
        token = self._issue_token(workbook_bytes)
        self._computed[token] = computed
        return ConversionResult.succeeded(
            token, summary, output_complete=output_complete
        )

    # ------------------------------------------------------------------
    # Optional invoice carton enrichment
    # ------------------------------------------------------------------
    def _attach_cartons(
        self, extracted: ExtractedDocument, invoice_raw: bytes
    ) -> ExtractedDocument:
        """Enrich line items from the supplier invoice / packing list (by serial).

        Parses the invoice line table and, for each ``LineItem`` whose serial
        appears there, sets the per-line carton count (``TOTAL CTNS`` -> Excel
        column ``G``) and overrides the description with the invoice's
        ``DESCRIPTION OF GOODS`` (Excel column ``E`` / the stock name), which the
        user prefers over the terser BOE description when an invoice is supplied.
        Lines absent from the invoice keep their BOE values. Any failure
        (unreadable invoice, unexpected layout) is swallowed and logged - the BOE
        conversion proceeds unchanged so an optional, malformed invoice never
        fails the whole conversion.

        For every matched line the BOE quantity, unit price and amount are also
        compared against the invoice's TOTAL QTY / UNIT PRICE / TOTAL AMOUNT;
        differing fields are recorded on the line (``invoice_mismatch``) so the
        Excel generator can red-fill that row. A field is only compared when
        both sides are numeric - anything missing or unreadable skips that
        field's comparison instead of flagging it.
        """
        try:
            import io

            details = self.invoice_parser.parse_line_details(io.BytesIO(invoice_raw))
        except Exception:
            logger.exception("Invoice extraction failed; proceeding without invoice data")
            return extracted

        if not details:
            return extracted

        new_items = []
        for item in extracted.line_items:
            detail = details.get(item.item_serial)
            if not detail:
                new_items.append(item)
                continue
            changes = {}
            if detail.get("cartons") is not None:
                changes["cartons"] = detail["cartons"]
            # When the invoice provides a description for this line, it takes
            # precedence over the BOE description (Excel column E / stock name).
            if detail.get("description") is not None:
                changes["description"] = detail["description"]
            mismatched = self._invoice_value_mismatches(item, detail)
            if mismatched:
                changes["invoice_mismatch"] = mismatched
            new_items.append(replace(item, **changes) if changes else item)
        return replace(extracted, line_items=new_items)

    @staticmethod
    def _invoice_value_mismatches(
        item: LineItem, detail: dict
    ) -> tuple[str, ...]:
        """Names of BOE-vs-invoice value fields that differ for one line.

        Compares quantity, unit price, and amount (the BOE amount is
        ``unit_price x quantity`` recomputed here; the BOE prints no per-line
        amount field). Quantity and unit price count only when both sides are
        numeric and differ beyond float dust (``math.isclose`` with a sub-paisa
        absolute tolerance). Amount is compared rounded to cents: the invoice
        prints amounts to 2 decimals, so a half-cent rounding difference
        (e.g. ``8333.5 x 0.15 = 1250.025`` printed as ``1250.03``) is not a
        mismatch - only a difference beyond half a cent plus dust flags.
        Returns ``()`` when nothing differs.
        """
        boe_qty = _as_number(item.quantity)
        boe_price = _as_number(item.unit_price_usd)
        boe_amount = (
            boe_price * boe_qty
            if boe_price is not None and boe_qty is not None
            else None
        )
        mismatched: list[str] = []
        inv_qty = _detail_number(detail, "quantity")
        if (
            boe_qty is not None
            and inv_qty is not None
            and not math.isclose(boe_qty, inv_qty, rel_tol=1e-9, abs_tol=1e-6)
        ):
            mismatched.append("quantity")
        inv_price = _detail_number(detail, "unit_price")
        if (
            boe_price is not None
            and inv_price is not None
            and not math.isclose(boe_price, inv_price, rel_tol=1e-9, abs_tol=1e-6)
        ):
            mismatched.append("unit_price_usd")
        inv_amount = _detail_number(detail, "amount")
        if boe_amount is not None and inv_amount is not None and not math.isclose(
            boe_amount, inv_amount, rel_tol=1e-9, abs_tol=0.005 + 1e-6
        ):
            mismatched.append("amount_usd")
        return tuple(mismatched)

    # ------------------------------------------------------------------
    # Download token store
    # ------------------------------------------------------------------
    def _issue_token(self, workbook_bytes: bytes) -> str:
        """Store the workbook bytes under a fresh token and return the token."""
        token = secrets.token_urlsafe(32)
        self._downloads[token] = workbook_bytes
        return token

    def get_download(self, token: str) -> bytes | None:
        """Return the workbook bytes for ``token``, or ``None`` if unknown.

        The Web_Interface's ``GET /api/download/{token}`` uses this to stream the
        ``.xlsx`` (404 when ``None``).
        """
        return self._downloads.get(token)

    def get_computed(self, token: str) -> ComputedDocument | None:
        """Return the ``ComputedDocument`` for ``token``, or ``None`` if unknown.

        Enables the separate Tally-JSON export to reuse the in-memory computed
        result (grouped ledgers, per-line landed cost, IGST) without re-parsing.
        """
        return self._computed.get(token)

    # ------------------------------------------------------------------
    # Cross-checks (declared vs extracted/computed)
    # ------------------------------------------------------------------
    def _cross_check(
        self,
        extracted: ExtractedDocument,
        computed: ComputedDocument,
        verification: GrandTotalVerification | None = None,
    ) -> tuple[list[Discrepancy], bool]:
        """Run the four cross-checks; return discrepancies and completeness.

        ``output_complete`` is ``False`` only when an item-count mismatch means
        the output must not be presented as complete (Req 3.3). The invoice-total,
        recompute, and grand-total checks never affect completeness (the workbook
        is retained and downloadable; Req 7.6/9.4/11.8).

        The ``verification`` is the ``GrandTotalVerification`` already computed in
        :meth:`convert` (so the fills and the summary discrepancies share one
        result). When ``None`` it is computed here so the cross-checks stay
        self-contained.
        """
        discrepancies: list[Discrepancy] = []
        output_complete = True

        item_count_disc = self._check_item_count(extracted)
        if item_count_disc is not None:
            discrepancies.append(item_count_disc)
            output_complete = False  # Req 3.3: not presented as complete.

        invoice_total_disc = self._check_invoice_total(extracted, computed)
        if invoice_total_disc is not None:
            discrepancies.append(invoice_total_disc)

        discrepancies.extend(self._check_recompute(computed))

        # Req 11.6/11.9: grand-total mismatches surface in the summary but never
        # affect output_complete (Req 11.8: the workbook is always retained).
        if verification is None:
            verification = self._verify_grand_totals(extracted, computed)
        discrepancies.extend(self._check_grand_totals(verification))

        return discrepancies, output_complete

    @staticmethod
    def _check_item_count(extracted: ExtractedDocument) -> Discrepancy | None:
        """Req 3.2/3.3: extracted count must equal the BOE's declared count.

        Returns a ``Discrepancy(ITEM_COUNT)`` carrying both the declared and the
        extracted counts when they differ; ``None`` when they match or the BOE
        declared no count to check against.
        """
        declared = extracted.declared_item_count
        actual = len(extracted.line_items)
        if declared is None or declared == actual:
            return None
        return Discrepancy(
            kind="ITEM_COUNT",
            message=(
                f"The BOE declares {declared} line items but {actual} "
                f"were extracted."
            ),
            expected=declared,
            actual=actual,
        )

    def _check_invoice_total(
        self, extracted: ExtractedDocument, computed: ComputedDocument
    ) -> Discrepancy | None:
        """Req 7.6/7.7: compare the computed USD total to the declared total.

        The computed total is the sum of per-line USD amounts
        (``Totals.total_amount_usd``). The declared total is the header invoice
        amount printed in the BOE. When the declared total is absent/unreadable a
        "could not be verified" discrepancy is produced (Req 7.7); otherwise a
        difference of more than 0.01 USD yields an ``INVOICE_TOTAL`` discrepancy
        carrying both values (Req 7.6). Differences within tolerance => ``None``.
        """
        computed_total = computed.totals.total_amount_usd
        declared_total = _as_number(extracted.header.invoice_amount)

        if declared_total is None:
            # Req 7.7: declared total absent or unreadable.
            return Discrepancy(
                kind="INVOICE_TOTAL",
                message=MESSAGE_TOTAL_UNVERIFIABLE,
                expected=None,
                actual=computed_total,
            )

        if abs(computed_total - declared_total) > TOLERANCE:
            return Discrepancy(
                kind="INVOICE_TOTAL",
                message=(
                    f"The computed invoice total {computed_total} USD differs "
                    f"from the BOE declared total {declared_total} USD."
                ),
                expected=declared_total,
                actual=computed_total,
            )
        return None

    def _check_recompute(self, computed: ComputedDocument) -> list[Discrepancy]:
        """Req 9.4: report extracted values that disagree with their recompute.

        For each line the value extracted from the BOE for ``total_duty`` is
        compared against the duty recomputed from its related fields
        (``combined_duty = (BCD + SWS) + IGST``). A difference of more than 0.01
        in the value's unit yields a ``Discrepancy(RECOMPUTE)`` reporting both the
        extracted and the recomputed value. Lines lacking either value are
        skipped (their missing inputs are already surfaced as ``ReviewFlag``s).
        """
        discrepancies: list[Discrepancy] = []
        for line in computed.lines:
            extracted_duty = _as_number(line.source.total_duty)
            recomputed_duty = line.combined_duty
            if extracted_duty is None or recomputed_duty is None:
                continue
            if abs(extracted_duty - recomputed_duty) > TOLERANCE:
                serial = line.source.item_serial
                discrepancies.append(
                    Discrepancy(
                        kind="RECOMPUTE",
                        message=(
                            f"Line item {serial}: extracted total duty "
                            f"{extracted_duty} differs from the recomputed "
                            f"value {recomputed_duty}."
                        ),
                        expected=extracted_duty,
                        actual=recomputed_duty,
                    )
                )
        return discrepancies

    # ------------------------------------------------------------------
    # Grand-total verification (Req 11)
    # ------------------------------------------------------------------
    def _verify_grand_totals(
        self, extracted: ExtractedDocument, computed: ComputedDocument
    ) -> GrandTotalVerification:
        """Req 11.1/11.2/11.3: verify the three Totals_Row grand totals vs the BOE.

        Compares each Excel grand total against the matching declared Part I
        duty-summary value printed in the BOE header, producing one
        :class:`GrandTotalCheck` per column:

        - **P** -- ``Totals.total_customs_duty`` vs ``declared_bcd + declared_sws``
          (both parts must be present; if either is missing/unreadable the column
          is ``UNVERIFIED``).
        - **Q** -- ``Totals.total_igst`` vs ``declared_igst``.
        - **S** -- ``Totals.total_customs_duty + Totals.total_igst`` vs
          ``declared_total_duty``.

        Each declared value is read via :func:`_as_number`, so a missing or
        unreadable field yields ``declared=None`` and ``status="UNVERIFIED"``
        (Req 11.7 -- could-not-verify, no fill). Otherwise a difference greater
        than ``GRAND_TOTAL_TOLERANCE_INR`` (1.00 INR) is a ``MISMATCH``
        (Req 11.4); within tolerance is a ``MATCH`` (Req 11.5).
        """
        totals = computed.totals
        header = extracted.header

        # Column P: BCD + SWS. Both parts are required; if either is
        # missing/unreadable the whole declared BCD+SWS is treated as absent.
        declared_bcd = _as_number(header.declared_bcd)
        declared_sws = _as_number(header.declared_sws)
        if declared_bcd is None or declared_sws is None:
            declared_p: float | None = None
        else:
            declared_p = declared_bcd + declared_sws

        checks = [
            self._grand_total_check("P", totals.total_customs_duty, declared_p),
            self._grand_total_check(
                "Q", totals.total_igst, _as_number(header.declared_igst)
            ),
            self._grand_total_check(
                "S",
                totals.total_customs_duty + totals.total_igst,
                _as_number(header.declared_total_duty),
            ),
        ]
        return GrandTotalVerification(checks=checks)

    @staticmethod
    def _grand_total_check(
        column: str, excel_total: float, declared: float | None
    ) -> GrandTotalCheck:
        """Classify one column as MATCH/MISMATCH/UNVERIFIED (Req 11.4/11.5/11.7)."""
        if declared is None:
            status = "UNVERIFIED"
        elif abs(excel_total - declared) > GRAND_TOTAL_TOLERANCE_INR:
            status = "MISMATCH"
        else:
            status = "MATCH"
        return GrandTotalCheck(
            column=column,
            declared=declared,
            excel_total=excel_total,
            status=status,
        )

    def _check_grand_totals(
        self, verification: GrandTotalVerification
    ) -> list[Discrepancy]:
        """Req 11.6/11.9: emit a ``GRAND_TOTAL`` discrepancy per MISMATCH column.

        Each MISMATCH yields a :class:`Discrepancy` naming the affected Excel
        column (P/Q/S) and carrying the declared value as ``expected`` and the
        Excel grand total as ``actual`` so both surface in the
        ``ConversionSummary`` (Req 11.9). ``UNVERIFIED`` columns represent a
        could-not-verify condition (Req 11.7): they emit an informational
        ``GRAND_TOTAL`` discrepancy with ``expected=None`` but never a mismatch.
        ``MATCH`` columns produce nothing. This check never affects
        ``output_complete`` (Req 11.8).
        """
        column_label = {
            "P": "BCD+SWS",
            "Q": "IGST",
            "S": "total duty",
        }
        discrepancies: list[Discrepancy] = []
        for check in verification.checks:
            if check.status == "MISMATCH":
                label = column_label[check.column]
                discrepancies.append(
                    Discrepancy(
                        kind="GRAND_TOTAL",
                        message=(
                            f"Excel column {check.column} grand total "
                            f"{check.excel_total} differs from the BOE declared "
                            f"{label} {check.declared}."
                        ),
                        expected=check.declared,
                        actual=check.excel_total,
                    )
                )
            elif check.status == "UNVERIFIED":
                label = column_label[check.column]
                discrepancies.append(
                    Discrepancy(
                        kind="GRAND_TOTAL",
                        message=(
                            f"Excel column {check.column} grand total "
                            f"{check.excel_total} could not be verified against "
                            f"the BOE declared {label} (value missing or "
                            f"unreadable)."
                        ),
                        expected=None,
                        actual=check.excel_total,
                    )
                )
        return discrepancies

    # ------------------------------------------------------------------
    # Summary (Req 9.1)
    # ------------------------------------------------------------------
    def _build_summary(
        self,
        extracted: ExtractedDocument,
        computed: ComputedDocument,
        discrepancies: list[Discrepancy],
    ) -> ConversionSummary:
        """Assemble the ``ConversionSummary`` rendered by the Web_Interface.

        Reports the number of line items extracted, the computed total invoice
        amount in USD, and the count of fields flagged for review (the aggregate
        of every extraction and computation ``ReviewFlag``), alongside the
        discrepancies surfaced by the cross-checks (Req 9.1).
        """
        flags: list[ReviewFlag] = list(computed.flags)
        return ConversionSummary(
            line_items_extracted=len(extracted.line_items),
            declared_item_count=extracted.declared_item_count,
            total_invoice_amount_usd=computed.totals.total_amount_usd,
            declared_invoice_amount_usd=_as_number(extracted.header.invoice_amount),
            review_flag_count=len(flags),
            review_flags=flags,
            discrepancies=discrepancies,
        )

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _safe_close(handle) -> None:
        """Close the validated PDF handle, ignoring any close-time errors."""
        if handle is None:
            return
        try:
            handle.close()
        except Exception:
            pass
