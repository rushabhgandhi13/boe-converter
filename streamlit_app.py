"""Streamlit front-end for the Bill of Entry -> CTN Excel converter.

This is a thin UI over :class:`boe_converter.orchestrator.ConversionOrchestrator`
(the same engine the FastAPI service uses). It does no parsing/formatting of its
own -- it uploads a BOE PDF, runs the full validate -> parse -> compute ->
generate pipeline, shows the conversion summary, and offers the styled ``.xlsx``
for download.

Run locally:    streamlit run streamlit_app.py
Deploy:         push to GitHub and point Streamlit Community Cloud at this file.

NOTE: like the Milestone 1 FastAPI endpoint, this app is unauthenticated; deploy
it only where that is acceptable (e.g. a private/internal Streamlit space).
"""

from __future__ import annotations

import hmac
import inspect
import json
import os

import streamlit as st

from boe_converter.orchestrator import ConversionOrchestrator

st.set_page_config(page_title="Bill of Entry Converter", page_icon="📄", layout="centered")


def _check_password() -> bool:
    """Gate the app behind a shared password stored in Streamlit secrets.

    The expected password is read from ``st.secrets["app_password"]`` (set in the
    Streamlit Cloud dashboard or a local ``.streamlit/secrets.toml``), so it is
    never committed to the repo. Authentication is remembered for the session.
    Comparison uses :func:`hmac.compare_digest` to avoid timing leaks.
    """
    if st.session_state.get("auth_ok"):
        return True

    expected = st.secrets.get("app_password")
    if not expected:
        st.error(
            "This app is not configured. Set `app_password` in the app's "
            "Streamlit secrets (Manage app → Settings → Secrets)."
        )
        st.stop()

    st.title("🔒 Bill of Entry Converter")
    with st.form("login", clear_on_submit=True):
        entered = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in")

    if submitted:
        if hmac.compare_digest(str(entered), str(expected)):
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


if not _check_password():
    st.stop()

st.title("Bill of Entry → CTN Excel Converter")
st.caption(
    "Upload an ICEGATE Bill of Entry PDF and download a CTN workbook formatted "
    "exactly like the reference sheet."
)


@st.cache_resource
def _orchestrator() -> ConversionOrchestrator:
    """A single shared orchestrator (holds the in-memory download store).

    Streamlit Community Cloud runs on a slow shared CPU, so a large
    multi-hundred-page Bill of Entry can take well over the default 60-second
    budget to parse. The budget is raised here (overridable via the
    ``BOE_TIME_BUDGET_SECONDS`` env var / app secret) so those conversions
    complete instead of failing with a generic CONVERSION_FAILED.
    """
    budget = float(os.environ.get("BOE_TIME_BUDGET_SECONDS", "600"))
    return ConversionOrchestrator(time_budget_seconds=budget)


convert_tab, json_tab = st.tabs(["📄 Convert to Excel", "🧾 Generate Tally JSON"])

# ---------------------------------------------------------------------------
# Tab 1: BOE PDF -> CTN Excel (the existing, unchanged conversion flow)
# ---------------------------------------------------------------------------
with convert_tab:
    uploaded = st.file_uploader("Bill of Entry PDF", type=["pdf"])
    invoice_uploaded = st.file_uploader(
        "Invoice / Packing List PDF (optional)",
        type=["pdf"],
        help=(
            "Optional. If provided, per-line carton counts (CTN, column G) are read "
            "from the invoice's TOTAL CTNS column and matched to the BOE line items "
            "by serial number. Leave empty to keep the CTN column blank."
        ),
    )
    usd_rate = st.number_input(
        "USD rate", min_value=0.0, value=95.30, step=0.01, format="%.2f",
        help="The USD→INR conversion rate applied to every line item.",
    )

    if st.button("Convert", type="primary", disabled=uploaded is None):
        raw = uploaded.getvalue()
        invoice_raw = (
            invoice_uploaded.getvalue() if invoice_uploaded is not None else None
        )
        orchestrator = _orchestrator()

        # Resilience against a stale/partial deploy: only pass invoice_raw if the
        # running orchestrator build actually supports it. This prevents a hard
        # TypeError crash when Streamlit Cloud has reloaded a newer streamlit_app.py
        # but an older orchestrator.py (a reboot fully syncs them).
        supports_invoice = (
            "invoice_raw" in inspect.signature(orchestrator.convert).parameters
        )
        if invoice_raw is not None and not supports_invoice:
            st.warning(
                "The invoice carton feature isn't active on this running app yet. "
                "Reboot the app (Manage app → Reboot) to enable it; converting the "
                "Bill of Entry without cartons for now."
            )

        with st.spinner("Converting…"):
            if supports_invoice:
                result = orchestrator.convert(
                    raw, uploaded.name, float(usd_rate), invoice_raw=invoice_raw
                )
            else:
                result = orchestrator.convert(raw, uploaded.name, float(usd_rate))

        if not result.ok:
            st.error(f"{result.error_code}: {result.message}")
            st.stop()

        summary = result.summary
        workbook = _orchestrator().get_download(result.download_token)

        # Remember the token so the JSON tab can reuse the in-memory computed
        # result (the direct "Generate JSON" path) without re-parsing the BOE.
        st.session_state["last_token"] = result.download_token
        st.session_state["last_usd_rate"] = float(usd_rate)

        if result.output_complete:
            st.success(f"Converted {summary.line_items_extracted} line items.")
        else:
            st.warning(
                f"Converted with an item-count mismatch: extracted "
                f"{summary.line_items_extracted}, BOE declared "
                f"{summary.declared_item_count}. Review before use."
            )

        c1, c2, c3 = st.columns(3)
        c1.metric("Line items", summary.line_items_extracted)
        c2.metric("Invoice total (USD)", f"{summary.total_invoice_amount_usd:,.2f}")
        c3.metric("Fields flagged", summary.review_flag_count)

        st.download_button(
            "⬇ Download CTN workbook (.xlsx)",
            data=workbook,
            file_name="bill_of_entry.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.info("Excel looks good? Switch to the **Generate Tally JSON** tab.")

        if summary.discrepancies:
            with st.expander(f"Discrepancies ({len(summary.discrepancies)})"):
                st.table(
                    [
                        {
                            "Type": d.kind,
                            "Expected": d.expected,
                            "Actual": d.actual,
                            "Detail": d.message,
                        }
                        for d in summary.discrepancies
                    ]
                )

        if summary.review_flags:
            with st.expander(f"Review flags ({summary.review_flag_count})"):
                st.table(
                    [
                        {
                            "Scope": f.scope,
                            "Item": f.item_serial,
                            "Field": f.field_name,
                            "Reason": f.reason,
                        }
                        for f in summary.review_flags
                    ]
                )


# ---------------------------------------------------------------------------
# Tab 2: Excel/computed -> Tally Purchase-voucher JSON (separate, optional flow)
# ---------------------------------------------------------------------------
with json_tab:
    from boe_converter.excel_reader import ExcelReadError, read_workbook
    from boe_converter.tally_exporter import CompanyProfile, TallyExporter
    from boe_converter.tally_master import TallyMaster

    st.subheader("Generate a Tally import JSON")
    st.caption(
        "Builds a Purchase voucher (line items grouped by IGST rate) that Tally "
        "can import. Every ledger name is matched against your Tally master so "
        "the master's exact spelling is used."
    )

    # --- Tally master (source of truth for names) ---
    master_file = st.file_uploader(
        "Tally Master.json (UTF-16, exported from Tally)",
        type=["json"],
        key="master_upload",
        help=(
            "Your company's master export. Ledger/party names in the JSON are "
            "matched against this so Tally accepts the import."
        ),
    )

    # --- Company (buyer) identity: company-level settings, not per-BOE data ---
    with st.expander("Company (buyer) details", expanded=False):
        company = CompanyProfile(
            name=st.text_input("Company name", value="Gemini Unicom LLP"),
            gstin=st.text_input("Company GSTIN", value=""),
            state=st.text_input("State (place of supply)", value="Maharashtra"),
            pincode=st.text_input("Pincode", value=""),
        )

    # --- Data source: direct (in-memory) or manual Excel upload ---
    manual = st.toggle(
        "Upload an edited Excel instead of using the last conversion",
        value=False,
        help=(
            "Off: use the workbook from the last conversion in this session. "
            "On: upload a CTN Excel you downloaded and edited (e.g. filled the "
            "'AS PER TALLY NAME' column). Open+save it in Excel first so formula "
            "values are cached."
        ),
    )

    excel_file = None
    if manual:
        excel_file = st.file_uploader(
            "Edited CTN workbook (.xlsx)", type=["xlsx"], key="json_excel_upload"
        )

    def _load_computed_and_rate():
        """Return (ComputedDocument, usd_rate) from the chosen source, or (None, ..)."""
        if manual:
            if excel_file is None:
                st.info("Upload an edited CTN workbook to continue.")
                return None, 0.0
            try:
                doc = read_workbook(excel_file.getvalue())
            except ExcelReadError as exc:
                st.error(str(exc))
                return None, 0.0
            return doc, doc.header.usd_rate or st.session_state.get("last_usd_rate", 0.0)
        token = st.session_state.get("last_token")
        doc = _orchestrator().get_computed(token) if token else None
        if doc is None:
            st.info(
                "No conversion found in this session. Convert a Bill of Entry in "
                "the first tab, or switch on the manual upload option above."
            )
            return None, 0.0
        return doc, st.session_state.get("last_usd_rate", doc.header.usd_rate)

    if st.button("Generate JSON", type="primary"):
        if master_file is None:
            st.error("Upload your Tally Master.json first.")
            st.stop()
        try:
            master = TallyMaster.load_bytes(master_file.getvalue())
        except Exception:
            st.error(
                "Could not read Master.json. It must be the UTF-16 file Tally "
                "exports."
            )
            st.stop()

        computed, rate = _load_computed_and_rate()
        if computed is None:
            st.stop()

        exporter = TallyExporter(master, company)

        # Master validation: which referenced ledgers are missing?
        required = exporter.required_ledger_names(computed)
        missing = master.missing(required)
        if missing:
            st.warning(
                "These ledgers are not in your Tally master. Tally will reject "
                "the import until they are added:"
            )
            st.table([{"Missing ledger": n} for n in missing])
            updated = TallyMaster.load_bytes(master_file.getvalue())
            for name in missing:
                updated.add_ledger(name)
            st.download_button(
                "⬇ Download updated Master.json (adds the missing ledgers)",
                data=updated.to_bytes(),
                file_name="Master_updated.json",
                mime="application/json",
                help=(
                    "Import this into Tally first, then import the voucher JSON "
                    "below."
                ),
            )
            st.caption(
                "Review the added ledgers in Tally (group/GSTIN) before importing "
                "the voucher."
            )

        # Build and offer the voucher JSON regardless (so the user can inspect it).
        document = exporter.build(computed, float(rate))
        payload = json.dumps(document, ensure_ascii=True, indent=1).encode("utf-8")
        st.success("Tally Purchase-voucher JSON generated.")
        st.download_button(
            "⬇ Download Tally voucher JSON",
            data=payload,
            file_name="tally_purchase_voucher.json",
            mime="application/json",
            type="primary",
        )
        with st.expander("Ledgers used in this voucher"):
            st.table([{"Ledger": n} for n in required])
