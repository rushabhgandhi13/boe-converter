"""Build a Tally *Purchase voucher* JSON from a computed BOE document.

This turns a :class:`~boe_converter.models.ComputedDocument` (the same in-memory
result that drives the Excel workbook) into the JSON shape Tally imports - a
single ``tallymessage`` Purchase voucher whose line items are **grouped by IGST
rate** into purchase ledgers, each carrying its stock items as
``inventoryallocations`` plus the matching ``IGST Purchase``/``IGST Payable``
ledgers, a ``Custom Duty Payable`` ledger and the supplier *party* ledger.

Accounting model (reverse-engineered from the reference voucher and expressed
purely in terms of computed per-line fields):

- **Purchase ledger** (grouped by IGST rate), debit: sum of
  ``land_cost_excl_gst`` for its lines; each stock item's amount is that line's
  ``land_cost_excl_gst`` and its rate is ``purchase_rate_per_unit``.
- **IGST Purchase** (debit) and **IGST Payable** (credit): equal and opposite,
  each the sum of ``igst_amount`` for the group (zero-rate groups have none).
- **Party** ledger (supplier), credit: total ``purchase_inr`` (the CIF value
  paid to the supplier = USD invoice x rate).
- **Custom Duty Payable**, credit: total ``total_customs_duty``.

The identity ``sum(land_cost_excl_gst) == sum(purchase_inr) + sum(total_customs_duty)``
keeps the voucher balanced.

Names are never invented: every ledger name is resolved through
:class:`~boe_converter.tally_master.TallyMaster` so the master's canonical
spelling is used. Only pure Tally structural constants live here.
"""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field

from boe_converter.models import ComputedDocument, ComputedLine
from boe_converter.tally_master import TallyMaster

# ---------------------------------------------------------------------------
# Pure Tally structural constants (never printed on a BOE)
# ---------------------------------------------------------------------------
NOT_APPLICABLE = "\u0004 Not Applicable"
COST_CATEGORY = "Primary Cost Category"
GODOWN = "Main Location"
BATCH = "Primary Batch"
DEFAULT_ENTERED_BY = "boe-converter"
CUSTOM_DUTY_LEDGER = "Custom Duty Payable"
TAX_FREE_LEDGER = "Tax Free (Purchases)"


@dataclass(frozen=True)
class CompanyProfile:
    """Buyer/company identity - the Tally company the voucher is imported into.

    These are company-level settings (not per-BOE data). Defaults are placeholders;
    callers should supply the real company's values (typically confirmed once in
    the UI). GSTIN/state drive Tally's tax context.
    """

    name: str = "Gemini Unicom LLP"
    gstin: str = ""
    state: str = "Maharashtra"
    pincode: str = ""
    entered_by: str = DEFAULT_ENTERED_BY
    address_lines: tuple[str, ...] = ()


# ---------------------------------------------------------------------------
# Number / quantity formatting (match the reference voucher's string forms)
# ---------------------------------------------------------------------------
def _amt(x: float) -> str:
    """Format a monetary amount as a 2-decimal string (e.g. ``-282818.99``)."""
    return f"{x:.2f}"


def _qty(q: float, unit: str) -> str:
    """Format a quantity as `` 4451.00 KGS`` (leading space, 2 decimals)."""
    return f" {q:.2f} {unit}"


def _rate(r: float, unit: str) -> str:
    """Format a unit rate as ``63.54/KGS``."""
    return f"{r:.2f}/{unit}"


def _pct(rate_fraction: float) -> float:
    """Convert a stored IGST fraction (0.05) to a percent number (5.0)."""
    return round(rate_fraction * 100, 2)


def _pct_label(rate_fraction: float) -> str:
    """A human percent label: ``5`` for whole, ``2.5`` for fractional."""
    p = _pct(rate_fraction)
    return str(int(p)) if float(p).is_integer() else ("%g" % p)


def _guid() -> str:
    return f"{uuid.uuid4()}-{uuid.uuid4().hex[:8]}"


# ---------------------------------------------------------------------------
# Value extraction helpers
# ---------------------------------------------------------------------------
def _num(rv) -> float | None:
    """Numeric interpretation of a RawValue-like field, else None."""
    if rv is None or getattr(rv, "is_missing", False) or getattr(rv, "is_unparseable", False):
        return None
    parsed = getattr(rv, "parsed", None)
    if isinstance(parsed, bool):
        return None
    if isinstance(parsed, (int, float)):
        return float(parsed)
    return None


def _text(rv) -> str | None:
    """Text interpretation of a RawValue-like field, else None."""
    if rv is None or getattr(rv, "is_missing", False):
        return None
    parsed = getattr(rv, "parsed", None)
    if isinstance(parsed, str) and parsed.strip():
        return parsed.strip()
    raw = getattr(rv, "raw_text", None)
    return raw.strip() if isinstance(raw, str) and raw.strip() else None


def _line_igst_fraction(line: ComputedLine) -> float:
    """The IGST rate fraction for a line (0.0 when missing/unreadable)."""
    r = _num(line.source.igst_rate)
    return r if r is not None else 0.0


def _stock_name(line: ComputedLine) -> str:
    """Stock item name: the BOE description (Excel col D is filled by a human).

    When a mapped Tally name is later supplied via the Excel upload path it can
    override this; from the in-memory document the verbatim description is used.
    """
    return _text(line.source.description) or f"ITEM {line.source.item_serial}"


# ---------------------------------------------------------------------------
# Exporter
# ---------------------------------------------------------------------------
class TallyExporter:
    """Builds a Purchase voucher ``tallymessage`` document."""

    def __init__(self, master: TallyMaster, company: CompanyProfile | None = None) -> None:
        self.master = master
        self.company = company or CompanyProfile()

    # -- public API ---------------------------------------------------------
    def required_ledger_names(self, computed: ComputedDocument) -> list[str]:
        """Every ledger name this voucher will reference, canonicalised.

        Used by the UI to check the master and prompt for missing entries
        *before* generating the JSON.
        """
        names: list[str] = []
        party = self._party_name(computed)
        names.append(self.master.resolve(party))

        rates = _distinct_rates(computed)
        for r in rates:
            if r <= 0:
                names.append(self.master.resolve(TAX_FREE_LEDGER))
                continue
            names.append(self.master.match_rate_ledger("Factory Purchase Import", r).name)
            names.append(self.master.match_rate_ledger("IGST Purchase", r).name)
            names.append(self.master.match_rate_ledger("IGST Payable", r).name)
        names.append(self.master.resolve(CUSTOM_DUTY_LEDGER))
        # De-duplicate preserving order.
        seen: set[str] = set()
        out: list[str] = []
        for n in names:
            if n not in seen:
                seen.add(n)
                out.append(n)
        return out

    def build(self, computed: ComputedDocument, usd_rate: float) -> dict:
        """Build the full Tally import document for ``computed``."""
        party_name = self._party_name(computed)
        party_ledger = self.master.resolve(party_name)
        cost_centre = _text(computed.header.details) or ""
        usd_total = computed.totals.total_amount_usd
        party_total = sum(l.purchase_inr or 0.0 for l in computed.lines)
        duty_total = sum(l.total_customs_duty or 0.0 for l in computed.lines)

        narration = self._narration(cost_centre, usd_total, usd_rate)

        ledger_entries: list[dict] = []
        # 1) Party (supplier) ledger - credit, carries the bill reference.
        ledger_entries.append(
            self._party_entry(party_ledger, party_total, computed)
        )

        # 2) Purchase + IGST ledgers, grouped by IGST rate.
        for rate in _distinct_rates(computed):
            group = [l for l in computed.lines if _same_rate(_line_igst_fraction(l), rate)]
            if rate <= 0:
                ledger_entries.append(
                    self._tax_free_entry(group, cost_centre)
                )
                continue
            ledger_entries.append(
                self._purchase_entry(rate, group, cost_centre)
            )
            igst_sum = sum(l.igst_amount or 0.0 for l in group)
            if abs(igst_sum) > 0:
                ledger_entries.append(self._igst_purchase_entry(rate, igst_sum, cost_centre))
                ledger_entries.append(self._igst_payable_entry(rate, igst_sum, cost_centre))

        # 3) Custom Duty Payable - credit.
        ledger_entries.append(
            self._simple_credit(self.master.resolve(CUSTOM_DUTY_LEDGER), duty_total, cost_centre=None)
        )

        voucher = self._voucher_shell(computed, party_ledger, narration, cost_centre)
        voucher["allledgerentries"] = ledger_entries
        return {"tallymessage": [voucher]}

    # -- party --------------------------------------------------------------
    def _party_name(self, computed: ComputedDocument) -> str:
        return _text(computed.header.party_name) or "Unknown Supplier"

    def _party_entry(self, ledger: str, total: float, computed: ComputedDocument) -> dict:
        bill_ref = _text(computed.header.invoice_no) or _text(computed.header.be_no) or "Ref"
        return {
            "oldauditentryids": [{"metadata": True, "type": "Number"}, "-1"],
            "ledgername": ledger,
            "gstclass": NOT_APPLICABLE,
            "isdeemedpositive": False,
            "ledgerfromitem": False,
            "removezeroentries": False,
            "ispartyledger": True,
            "amount": _amt(total),
            "vatexpamount": _amt(total),
            "billallocations": [
                {
                    "name": bill_ref,
                    "billtype": "New Ref",
                    "tdsdeducteeisspecialrate": False,
                    "amount": _amt(total),
                }
            ],
        }

    # -- purchase (taxable, grouped) ---------------------------------------
    def _purchase_entry(self, rate: float, group: list[ComputedLine], cost_centre: str) -> dict:
        ledger = self.master.match_rate_ledger("Factory Purchase Import", rate).name
        total = sum(l.land_cost_excl_gst or 0.0 for l in group)
        return {
            "oldauditentryids": [{"metadata": True, "type": "Number"}, "-1"],
            "ledgername": ledger,
            "gstclass": NOT_APPLICABLE,
            "gstovrdnineligibleitc": NOT_APPLICABLE,
            "gstovrdnisrevchargeappl": NOT_APPLICABLE,
            "gstovrdntypeofsupply": "Goods",
            "gstrateinferapplicability": "As per Masters/Company",
            "gsthsninferapplicability": "As per Masters/Company",
            "isdeemedpositive": True,
            "ledgerfromitem": False,
            "removezeroentries": False,
            "ispartyledger": False,
            "islastdeemedpositive": True,
            "amount": _amt(-total),
            "vatexpamount": _amt(-total),
            "inventoryallocations": [
                self._inventory(l, rate, cost_centre) for l in group
            ],
        }

    def _inventory(self, line: ComputedLine, rate: float, cost_centre: str) -> dict:
        name = _stock_name(line)
        hsn = _text(line.source.cth_hsn) or ""
        unit = _text(line.source.unit) or "NOS"
        qty = _num(line.source.quantity) or 0.0
        amount = line.land_cost_excl_gst or 0.0
        unit_rate = line.purchase_rate_per_unit or (amount / qty if qty else 0.0)
        pct = _pct(rate)
        half = round(pct / 2, 2)
        alloc = {
            "stockitemname": name,
            "gstovrdnineligibleitc": NOT_APPLICABLE,
            "gstovrdnisrevchargeappl": NOT_APPLICABLE,
            "gstovrdntaxability": "Taxable",
            "gstsourcetype": "Stock Item",
            "gstitemsource": name,
            "hsnsourcetype": "Stock Item",
            "hsnitemsource": name,
            "gstovrdntypeofsupply": "Goods",
            "gstrateinferapplicability": "As per Masters/Company",
            "gsthsnname": hsn,
            "gsthsninferapplicability": "As per Masters/Company",
            "isdeemedpositive": True,
            "islastdeemedpositive": True,
            "isautonegate": False,
            "iscustomsclearance": False,
            "rate": _rate(unit_rate, unit),
            "amount": _amt(-amount),
            "actualqty": _qty(qty, unit),
            "billedqty": _qty(qty, unit),
            "categoryallocations": [
                {
                    "category": COST_CATEGORY,
                    "isdeemedpositive": True,
                    "costcentreallocations": [
                        {"name": cost_centre, "amount": _amt(-amount)}
                    ],
                }
            ],
            "batchallocations": [
                {
                    "godownname": GODOWN,
                    "batchname": BATCH,
                    "indentno": NOT_APPLICABLE,
                    "orderno": NOT_APPLICABLE,
                    "trackingnumber": NOT_APPLICABLE,
                    "dynamiccstiscleared": False,
                    "amount": _amt(-amount),
                    "actualqty": _qty(qty, unit),
                    "billedqty": _qty(qty, unit),
                }
            ],
            "ratedetails": [
                {"gstratedutyhead": "CGST", "gstratevaluationtype": "Based on Value", "gstrate": f" {half}"},
                {"gstratedutyhead": "SGST/UTGST", "gstratevaluationtype": "Based on Value", "gstrate": f" {half}"},
                {"gstratedutyhead": "IGST", "gstratevaluationtype": "Based on Value", "gstrate": f" {_pct_label(rate)}"},
                {"gstratedutyhead": "Cess", "gstratevaluationtype": NOT_APPLICABLE},
                {"gstratedutyhead": "State Cess", "gstratevaluationtype": "Based on Value"},
            ],
        }
        return alloc

    # -- tax free (zero-rated) ---------------------------------------------
    def _tax_free_entry(self, group: list[ComputedLine], cost_centre: str) -> dict:
        ledger = self.master.resolve(TAX_FREE_LEDGER)
        total = sum(l.land_cost_excl_gst or 0.0 for l in group)
        return {
            "oldauditentryids": [{"metadata": True, "type": "Number"}, "-1"],
            "ledgername": ledger,
            "gstclass": NOT_APPLICABLE,
            "gstovrdnineligibleitc": NOT_APPLICABLE,
            "gstovrdnisrevchargeappl": NOT_APPLICABLE,
            "gstovrdntaxability": "Nil Rated",
            "gstsourcetype": "Ledger",
            "gstledgersource": ledger,
            "gstovrdntypeofsupply": "Services",
            "isdeemedpositive": True,
            "ledgerfromitem": False,
            "removezeroentries": False,
            "ispartyledger": False,
            "islastdeemedpositive": True,
            "amount": _amt(-total),
            "vatexpamount": _amt(-total),
            "inventoryallocations": [
                self._inventory_nil(l, cost_centre) for l in group
            ],
        }

    def _inventory_nil(self, line: ComputedLine, cost_centre: str) -> dict:
        name = _stock_name(line)
        hsn = _text(line.source.cth_hsn) or ""
        unit = _text(line.source.unit) or "NOS"
        qty = _num(line.source.quantity) or 0.0
        amount = line.land_cost_excl_gst or 0.0
        unit_rate = line.purchase_rate_per_unit or (amount / qty if qty else 0.0)
        return {
            "stockitemname": name,
            "gstovrdntaxability": "Nil Rated",
            "gstsourcetype": "Stock Item",
            "gstitemsource": name,
            "hsnsourcetype": "Stock Item",
            "hsnitemsource": name,
            "gstovrdntypeofsupply": "Goods",
            "gsthsnname": hsn,
            "isdeemedpositive": True,
            "islastdeemedpositive": True,
            "rate": _rate(unit_rate, unit),
            "amount": _amt(-amount),
            "actualqty": _qty(qty, unit),
            "billedqty": _qty(qty, unit),
            "categoryallocations": [
                {
                    "category": COST_CATEGORY,
                    "isdeemedpositive": True,
                    "costcentreallocations": [
                        {"name": cost_centre, "amount": _amt(-amount)}
                    ],
                }
            ],
            "batchallocations": [
                {
                    "godownname": GODOWN,
                    "batchname": BATCH,
                    "amount": _amt(-amount),
                    "actualqty": _qty(qty, unit),
                    "billedqty": _qty(qty, unit),
                }
            ],
        }

    # -- IGST purchase / payable -------------------------------------------
    def _igst_purchase_entry(self, rate: float, igst: float, cost_centre: str) -> dict:
        ledger = self.master.match_rate_ledger("IGST Purchase", rate).name
        return self._tax_entry(ledger, -igst, deemed_positive=True, cost_centre=cost_centre)

    def _igst_payable_entry(self, rate: float, igst: float, cost_centre: str) -> dict:
        ledger = self.master.match_rate_ledger("IGST Payable", rate).name
        return self._tax_entry(ledger, igst, deemed_positive=False, cost_centre=cost_centre)

    def _tax_entry(self, ledger: str, amount: float, deemed_positive: bool, cost_centre: str) -> dict:
        return {
            "oldauditentryids": [{"metadata": True, "type": "Number"}, "-1"],
            "ledgername": ledger,
            "gstclass": NOT_APPLICABLE,
            "isdeemedpositive": deemed_positive,
            "ledgerfromitem": False,
            "removezeroentries": False,
            "ispartyledger": False,
            "islastdeemedpositive": deemed_positive,
            "amount": _amt(amount),
            "vatexpamount": _amt(amount),
            "categoryallocations": [
                {
                    "category": COST_CATEGORY,
                    "isdeemedpositive": deemed_positive,
                    "costcentreallocations": [
                        {"name": cost_centre, "amount": _amt(amount)}
                    ],
                }
            ],
        }

    def _simple_credit(self, ledger: str, amount: float, cost_centre: str | None) -> dict:
        return {
            "oldauditentryids": [{"metadata": True, "type": "Number"}, "-1"],
            "ledgername": ledger,
            "gstclass": NOT_APPLICABLE,
            "isdeemedpositive": False,
            "ledgerfromitem": False,
            "removezeroentries": False,
            "ispartyledger": False,
            "islastdeemedpositive": False,
            "amount": _amt(amount),
            "vatexpamount": _amt(amount),
        }

    # -- voucher shell ------------------------------------------------------
    def _narration(self, cost_centre: str, usd_total: float, usd_rate: float) -> str:
        base = cost_centre or ""
        return f"{base} USD {usd_total:.2f} @{_fmt_num(usd_rate)}".strip()

    def _voucher_shell(
        self,
        computed: ComputedDocument,
        party_ledger: str,
        narration: str,
        cost_centre: str,
    ) -> dict:
        date = _tally_date(_text(computed.header.be_date))
        guid = _guid()
        party_name = self._party_name(computed)
        return {
            "metadata": {
                "type": "Voucher",
                "guid": guid,
                "vchtype": "Purchase",
                "action": "Create",
                "objview": "Accounting Voucher View",
            },
            "date": date,
            "referencedate": date,
            "effectivedate": date,
            "guid": guid,
            "vatdealertype": "Regular",
            "narration": narration,
            "enteredby": self.company.entered_by,
            "vouchertypename": "Purchase",
            "partyname": party_name,
            "partyledgername": party_ledger,
            "partymailingname": party_name,
            "basicbasepartyname": party_name,
            "basicbuyername": self.company.name,
            "placeofsupply": self.company.state,
            "cmpgststate": self.company.state,
            "consigneestatename": self.company.state,
            "consigneecountryname": "India",
            "cmpgstin": self.company.gstin,
            "consigneegstin": self.company.gstin,
            "consigneepincode": self.company.pincode,
            "reference": _text(computed.header.invoice_no) or "",
            "costcentrename": cost_centre,
            "numberingstyle": "Auto",
            "vchentrymode": "As Voucher",
            "iscostcentre": True,
            "persistedview": "Accounting Voucher View",
            "isinvoice": False,
            "isdeemedpositive": False,
            "iseligibleforitc": True,
        }


# ---------------------------------------------------------------------------
# Module-level helpers
# ---------------------------------------------------------------------------
def _fmt_num(x: float) -> str:
    return str(int(x)) if float(x).is_integer() else ("%g" % x)


def _same_rate(a: float, b: float) -> bool:
    return round(a * 10000) == round(b * 10000)


def _distinct_rates(computed: ComputedDocument) -> list[float]:
    """Distinct IGST rate fractions present, ascending."""
    rates: list[float] = []
    for line in computed.lines:
        r = _line_igst_fraction(line)
        if not any(_same_rate(r, x) for x in rates):
            rates.append(r)
    return sorted(rates)


def _tally_date(be_date: str | None) -> str:
    """Convert a BOE date (dd/mm/yyyy or dd-mm-yyyy) to Tally ``YYYYMMDD``.

    Falls back to an empty string when unparseable (Tally then uses the import
    date); never guesses a wrong date.
    """
    if not be_date:
        return ""
    import re

    m = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", be_date)
    if not m:
        return ""
    d, mo, y = m.groups()
    if len(y) == 2:
        y = "20" + y
    return f"{int(y):04d}{int(mo):02d}{int(d):02d}"
