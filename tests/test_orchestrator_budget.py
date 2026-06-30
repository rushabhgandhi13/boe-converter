"""Tests for the configurable post-recognition time budget (Req 1.3/1.7).

The orchestrator's time budget defaults to 60 seconds but is configurable so a
slow deployment (e.g. a shared-CPU Streamlit host converting a large
multi-hundred-page Bill of Entry) can allow more time instead of failing a
conversion that would otherwise succeed. A genuine overrun is still a
post-recognition failure that issues no download token (atomic output).

These tests use lightweight stubs (no real PDF) so they are fast and
deterministic: a stub validator reports the document as a recognized BOE, and a
deliberately slow parser forces the elapsed time past a tiny budget.
"""

from __future__ import annotations

import time

from boe_converter.orchestrator import (
    ConversionOrchestrator,
    ERROR_CONVERSION_FAILED,
)
from boe_converter.validator import ValidationOutcome


class _OkValidator:
    """A validator stub that always accepts (recognized BOE), no real PDF."""

    def validate(self, raw, filename):  # noqa: ANN001
        return ValidationOutcome.success(handle=None)


class _SlowParser:
    """A parser stub that sleeps to push the pipeline past a tiny budget."""

    def __init__(self, delay: float) -> None:
        self._delay = delay

    def parse(self, handle, *, usd_rate=0.0):  # noqa: ANN001
        time.sleep(self._delay)
        # The result is irrelevant: the budget check fires before success.
        from boe_converter.models import ExtractedDocument, HeaderBlock, RawValue

        header = HeaderBlock(
            company_name="X", party_name=RawValue.missing(), usd_rate=usd_rate,
            details=RawValue.missing(), invoice_no=RawValue.missing(),
            invoice_date=RawValue.missing(), be_no=RawValue.missing(),
            be_date=RawValue.missing(), bl_no=RawValue.missing(),
            bl_date=RawValue.missing(), invoice_amount=RawValue.missing(),
            invoice_currency=RawValue.missing(), package_count=RawValue.missing(),
            container_details=RawValue.missing(),
        )
        return ExtractedDocument(header=header, line_items=[], declared_item_count=0)


def test_default_budget_is_60_seconds():
    orch = ConversionOrchestrator()
    assert orch.time_budget_seconds == 60.0


def test_budget_is_configurable():
    orch = ConversionOrchestrator(time_budget_seconds=600)
    assert orch.time_budget_seconds == 600


def test_overrun_fails_with_no_token():
    """A conversion that exceeds the (tiny) budget is a post-recognition
    failure: CONVERSION_FAILED, no token, nothing retrievable."""
    orch = ConversionOrchestrator(
        validator=_OkValidator(),
        parser=_SlowParser(delay=0.05),
        time_budget_seconds=0.0,  # any post-recognition work overruns
    )
    result = orch.convert(b"%PDF-fake", "boe.pdf", 95.3)

    assert result.ok is False
    assert result.error_code == ERROR_CONVERSION_FAILED
    assert result.download_token is None
    assert orch._downloads == {}


def test_generous_budget_allows_a_slow_conversion():
    """The same slow conversion succeeds when the budget is generous."""
    orch = ConversionOrchestrator(
        validator=_OkValidator(),
        parser=_SlowParser(delay=0.05),
        time_budget_seconds=30.0,
    )
    result = orch.convert(b"%PDF-fake", "boe.pdf", 95.3)

    assert result.ok is True
    assert result.download_token is not None
    assert orch.get_download(result.download_token) is not None
