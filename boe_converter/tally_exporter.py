"""Build a Tally *Purchase voucher* JSON from a computed BOE document.

This turns a :class:`~boe_converter.models.ComputedDocument` (the same in-memory
result that drives the Excel workbook) into the JSON shape Tally imports - a
single ``tallymessage`` Purchase voucher whose line items are **grouped by IGST
rate** into purchase ledgers, each carrying its stock items as
``inventoryallocations`` plus the matching ``IGST Purchase``/``IGST Payable``
ledgers, a ``Custom Duty Payable`` ledger and the supplier *party* ledger.

Buyer (importer) and seller (supplier) identity are populated **from the BOE**
(the ``HeaderBlock`` buyer_*/seller_* fields the parser now extracts), with an
optional :class:`CompanyProfile` / :class:`SellerProfile` override for any field
the user wants to correct from stored data. No Tally master file is consulted:
ledger names follow Tally's own deterministic conventions
(``Factory Purchase (Import 5%)``, ``IGST Purchase @ 5.00 %``,
``IGST Payable @ 5%``, ``Custom Duty Payable``, ``Tax Free (Purchases)``).

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

Unit conversion: quantities printed in dozens/gross/thousand are converted to
pieces on the inventory allocation (DOZ x12, GRS x144, THD x1000) so Tally
receives PCS. The line amount is preserved; only qty/unit/rate change.
"""

from __future__ import annotations

import uuid
from dataclasses import dataclass, replace

from boe_converter.models import ComputedDocument, ComputedLine, HeaderBlock, RawValue

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

# Unit conversion factors to pieces (PCS). A unit not listed is left unchanged.
_UNIT_TO_PCS = {"DOZ": 12.0, "GRS": 144.0, "THD": 1000.0}


@dataclass(frozen=True)
class CompanyProfile:
    """Buyer/company identity - the Tally company the voucher is imported into.

    Every field defaults to ``None`` meaning "use the value extracted from the
    BOE". Supply a value only to override the BOE (e.g. from stored data). This
    is how the buyer is populated *from the document* while still allowing a
    manual correction.
    """

    name: str | None = None
    gstin: str | None = None
    state: str | None = None
    pincode: str | None = None
    address_lines: tuple[str, ...] | None = None
    entered_by: str = DEFAULT_ENTERED_BY


@dataclass(frozen=True)
class SellerProfile:
    """Supplier/seller identity override (defaults to the BOE-extracted values)."""

    name: str | None = None
    address_lines: tuple[str, ...] | None = None
    country: str | None = None


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


def _convert_to_pcs(qty: float, unit: str) -> tuple[float, str]:
    """Convert a (qty, unit) to pieces when the unit is DOZ/GRS/THD.

    Returns ``(converted_qty, converted_unit)``; unchanged for other units.
    """
    factor = _UNIT_TO_PCS.get(unit.upper().strip()) if unit else None
    if factor:
        return qty * factor, "PCS"
    return qty, unit


# ---------------------------------------------------------------------------
# Ledger-name conventions (deterministic; no master file needed)
# ---------------------------------------------------------------------------
def _purchase_ledger_name(rate_fraction: float) -> str:
    return f"Factory Purchase (Import {_pct_label(rate_fraction)}%)"


def _igst_purchase_ledger_name(rate_fraction: float) -> str:
    return f"IGST Purchase @ {_pct(rate_fraction):.2f} %"


def _igst_payable_ledger_name(rate_fraction: float) -> str:
    return f"IGST Payable @ {_pct_label(rate_fraction)}%"


def _party_ledger_name(supplier: str) -> str:
    """Tally party ledger name: title-cased supplier name (matches reference)."""
    return supplier.title()


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


def _split_address(text: str | None) -> list[str]:
    """Split a comma-joined address into individual lines (empty when absent)."""
    if not text:
        return []
    return [part.strip() for part in text.split(",") if part.strip()]


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
    """Builds a Purchase voucher ``tallymessage`` document.

    Buyer/seller identity comes from the BOE ``HeaderBlock`` by default; pass a
    :class:`CompanyProfile` / :class:`SellerProfile` to override individual
    fields (e.g. from stored data).
    """

    def __init__(
        self,
        company: CompanyProfile | None = None,
        seller: SellerProfile | None = None,
    ) -> None:
        self.company = company or CompanyProfile()
        self.seller = seller or SellerProfile()

    # -- public API ---------------------------------------------------------
    def required_ledger_names(self, computed: ComputedDocument) -> list[str]:
        """Every ledger name this voucher will reference (for UI display)."""
        names: list[str] = [self._party_name(computed)]
        for r in _distinct_rates(computed):
            if r <= 0:
                names.append(TAX_FREE_LEDGER)
                continue
            names.append(_purchase_ledger_name(r))
            names.append(_igst_purchase_ledger_name(r))
            names.append(_igst_payable_ledger_name(r))
        names.append(CUSTOM_DUTY_LEDGER)
        seen: set[str] = set()
        out: list[str] = []
        for n in names:
            if n not in seen:
                seen.add(n)
                out.append(n)
        return out

    def build(self, computed: ComputedDocument, usd_rate: float) -> dict:
        """Build the full Tally import document for ``computed``."""
        party_ledger = self._party_ledger(computed)
        cost_centre = _text(computed.header.details) or ""
        usd_total = computed.totals.total_amount_usd
        party_total = sum(l.purchase_inr or 0.0 for l in computed.lines)
        duty_total = sum(l.total_customs_duty or 0.0 for l in computed.lines)

        narration = self._narration(cost_centre, usd_total, usd_rate)

        ledger_entries: list[dict] = []
        # 1) Party (supplier) ledger - credit, carries the bill reference.
        ledger_entries.append(self._party_entry(party_ledger, party_total, computed))

        # 2) Purchase + IGST ledgers, grouped by IGST rate.
        for rate in _distinct_rates(computed):
            group = [l for l in computed.lines if _same_rate(_line_igst_fraction(l), rate)]
            if rate <= 0:
                ledger_entries.append(self._tax_free_entry(group, cost_centre))
                continue
            ledger_entries.append(self._purchase_entry(rate, group, cost_centre))
            igst_sum = sum(l.igst_amount or 0.0 for l in group)
            if abs(igst_sum) > 0:
                ledger_entries.append(self._igst_purchase_entry(rate, igst_sum, cost_centre))
                ledger_entries.append(self._igst_payable_entry(rate, igst_sum, cost_centre))

        # 3) Custom Duty Payable - credit.
        ledger_entries.append(
            self._simple_credit(CUSTOM_DUTY_LEDGER, duty_total, cost_centre=None)
        )

        voucher = self._voucher_shell(computed, party_ledger, narration, cost_centre)
        voucher["allledgerentries"] = ledger_entries
        return {"tallymessage": [voucher]}

    # -- buyer / seller resolution -----------------------------------------
    def _buyer_name(self, computed: ComputedDocument) -> str:
        if self.company.name:
            return self.company.name
        name = computed.header.company_name or ""
        # The BOE importer name carries an "M/S " prefix for the Excel sheet;
        # strip it for the Tally buyer so it matches the company master.
        if name[:4].upper() == "M/S ":
            name = name[4:].strip()
        return name

    def _buyer_gstin(self, computed: ComputedDocument) -> str:
        return self.company.gstin or _text(computed.header.buyer_gstin) or ""

    def _buyer_state(self, computed: ComputedDocument) -> str:
        return self.company.state or _text(computed.header.buyer_state) or ""

    def _buyer_pincode(self, computed: ComputedDocument) -> str:
        return self.company.pincode or _text(computed.header.buyer_pincode) or ""

    def _buyer_address_lines(self, computed: ComputedDocument) -> list[str]:
        if self.company.address_lines is not None:
            return list(self.company.address_lines)
        return _split_address(_text(computed.header.buyer_address))

    def _seller_address_lines(self, computed: ComputedDocument) -> list[str]:
        if self.seller.address_lines is not None:
            return list(self.seller.address_lines)
        return _split_address(_text(computed.header.seller_address))

    def _seller_country(self, computed: ComputedDocument) -> str:
        return self.seller.country or _text(computed.header.seller_country) or ""

    # -- party --------------------------------------------------------------
    def _party_name(self, computed: ComputedDocument) -> str:
        return self.seller.name or _text(computed.header.party_name) or "Unknown Supplier"

    def _party_ledger(self, computed: ComputedDocument) -> str:
        return _party_ledger_name(self._party_name(computed))

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
        ledger = _purchase_ledger_name(rate)
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
        raw_unit = _text(line.source.unit) or "NOS"
        raw_qty = _num(line.source.quantity) or 0.0
        amount = line.land_cost_excl_gst or 0.0
        # Convert dozens/gross/thousand to pieces; keep the amount, adjust rate.
        qty, unit = _convert_to_pcs(raw_qty, raw_unit)
        unit_rate = (amount / qty) if qty else 0.0
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
        ledger = TAX_FREE_LEDGER
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
        raw_unit = _text(line.source.unit) or "NOS"
        raw_qty = _num(line.source.quantity) or 0.0
        amount = line.land_cost_excl_gst or 0.0
        qty, unit = _convert_to_pcs(raw_qty, raw_unit)
        unit_rate = (amount / qty) if qty else 0.0
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
        ledger = _igst_purchase_ledger_name(rate)
        return self._tax_entry(ledger, -igst, deemed_positive=True, cost_centre=cost_centre)

    def _igst_payable_entry(self, rate: float, igst: float, cost_centre: str) -> dict:
        ledger = _igst_payable_ledger_name(rate)
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

    def _string_block(self, lines: list[str]) -> list:
        """A Tally multi-line string block: ``[{metadata..}, line, line, ...]``."""
        return [{"metadata": True, "type": "String"}, *lines]

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
        buyer_name = self._buyer_name(computed)
        buyer_gstin = self._buyer_gstin(computed)
        buyer_state = self._buyer_state(computed)
        buyer_pincode = self._buyer_pincode(computed)
        buyer_addr = self._buyer_address_lines(computed)
        seller_addr = self._seller_address_lines(computed)
        seller_country = self._seller_country(computed)

        shell = {
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
            "countryofresidence": seller_country or "India",
            "vouchertypename": "Purchase",
            "partyname": party_name,
            "partyledgername": party_ledger,
            "partymailingname": party_name,
            "basicbasepartyname": party_name,
            "basicbuyername": buyer_name,
            "placeofsupply": buyer_state,
            "cmpgststate": buyer_state,
            "consigneestatename": buyer_state,
            "consigneecountryname": "India",
            "cmpgstin": buyer_gstin,
            "consigneegstin": buyer_gstin,
            "consigneepincode": buyer_pincode,
            "consigneemailingname": buyer_name,
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
        # Seller (supplier) address block and buyer address block, populated
        # from the BOE so Tally shows the correct parties.
        if seller_addr:
            shell["address"] = self._string_block(seller_addr)
        if buyer_addr:
            shell["basicbuyeraddress"] = self._string_block(buyer_addr)
        return shell


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


def apply_stock_names(computed: ComputedDocument, names: dict[int, str]) -> ComputedDocument:
    """Return a copy of ``computed`` with mapped "as per Tally" stock names.

    ``names`` maps a line's serial number to the canonical Tally stock-item name
    chosen in Step 2. The name replaces that line's ``description`` (the field
    the exporter reads for ``stockitemname``); lines without a mapping are left
    unchanged. This lets the in-memory JSON path use the Step 2 selections
    without going through an Excel round-trip.
    """
    if not names:
        return computed
    new_lines = []
    for line in computed.lines:
        mapped = names.get(line.source.item_serial)
        if mapped and mapped.strip():
            new_source = replace(
                line.source,
                description=RawValue(raw_text=mapped.strip(), parsed=mapped.strip()),
            )
            new_lines.append(replace(line, source=new_source))
        else:
            new_lines.append(line)
    return replace(computed, lines=new_lines)


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
