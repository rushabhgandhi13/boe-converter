"""FastAPI application factory for the BOE converter Web_Interface.

Implements task 11.1: the HTTP API and the single-file upload page for
Milestone 1.

Endpoints (design.md "Web_Interface (HTTP API)"):

- ``POST /api/convert`` -- accepts ``multipart/form-data`` with exactly one
  ``file`` (the BOE PDF) and a ``usd_rate`` number. On success returns
  ``{ "download_token": str, "summary": ConversionSummary }`` (200). On a
  rejection or post-recognition failure it returns the mapped error body
  ``{ "error_code": str, "message": str }`` with the status code mapping below.
- ``GET /api/download/{download_token}`` -- streams the generated ``.xlsx``
  bytes (200) or 404 when the token is unknown/expired.
- ``GET /`` -- serves the single-file upload page with a processing indicator
  and a results/summary view that enables the download.

Status-code mapping (design "Error Handling" + Web_Interface contract):

    FILE_TOO_LARGE     -> 413  (size rejection, Req 1.2)
    INVALID_PDF        -> 400  (unreadable PDF rejection, Req 1.5)
    NOT_A_BOE          -> 422  (not a recognized BOE, Req 1.6)
    CONVERSION_FAILED  -> 422  (post-recognition failure, Req 1.7)
    invalid/missing usd_rate -> 422 (request validation)

SECURITY NOTE (Milestone 1): these endpoints are UNAUTHENTICATED. The converter
is intended to run as a local, single-user tool for Milestone 1, so there is no
authentication or access control on upload/download. Do not expose this app on
an untrusted network without adding auth.
"""

from __future__ import annotations

from dataclasses import asdict

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, Response

from boe_converter.models import ConversionSummary
from boe_converter.orchestrator import (
    ConversionOrchestrator,
    ERROR_CONVERSION_FAILED,
)
from boe_converter.validator import (
    ERROR_FILE_TOO_LARGE,
    ERROR_INVALID_PDF,
    ERROR_NOT_A_BOE,
)

# The xlsx content type for the download response (design Web_Interface).
XLSX_MEDIA_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Error code -> HTTP status mapping (design "Error Handling").
_ERROR_STATUS: dict[str, int] = {
    ERROR_FILE_TOO_LARGE: 413,
    ERROR_INVALID_PDF: 400,
    ERROR_NOT_A_BOE: 422,
    ERROR_CONVERSION_FAILED: 422,
}

# Error code for an invalid/missing usd_rate form field (request validation).
ERROR_INVALID_USD_RATE = "INVALID_USD_RATE"
MESSAGE_INVALID_USD_RATE = "A numeric USD rate is required."


def _summary_to_dict(summary: ConversionSummary) -> dict:
    """Serialize a :class:`ConversionSummary` (a dataclass) to a JSON-able dict.

    ``dataclasses.asdict`` recurses into the nested ``review_flags`` and
    ``discrepancies`` dataclass lists, producing plain dicts the JSON response
    encoder can handle.
    """
    return asdict(summary)


def _error_response(error_code: str, message: str) -> JSONResponse:
    """Build the mapped ``{ error_code, message }`` body with its status code."""
    status = _ERROR_STATUS.get(error_code, 422)
    return JSONResponse(
        status_code=status,
        content={"error_code": error_code, "message": message},
    )


def create_app() -> FastAPI:
    """Create and configure the FastAPI app for the BOE Web_Interface.

    Builds a single shared :class:`ConversionOrchestrator` (it owns the
    in-memory download-token store), registers the two API routes plus the
    ``GET /`` upload page, and returns the configured app.
    """
    app = FastAPI(title="Bill of Entry Converter", version="0.1.0")

    # Shared orchestrator instance: holds the in-memory token -> workbook store
    # so a token issued by /api/convert is retrievable by /api/download.
    orchestrator = ConversionOrchestrator()

    @app.get("/", response_class=HTMLResponse)
    def index() -> HTMLResponse:
        """Serve the single-file upload page (processing + results view)."""
        return HTMLResponse(content=_UPLOAD_PAGE_HTML)

    @app.post("/api/convert")
    async def convert(
        # NOTE (Milestone 1): unauthenticated endpoint -- local tool only.
        file: UploadFile = File(...),
        usd_rate: str = Form(...),
        invoice: UploadFile | None = File(None),
    ):
        """Convert one uploaded BOE PDF into the CTN workbook.

        An optional ``invoice`` PDF (the supplier invoice / packing list) may be
        supplied; when present its per-line carton counts are read and matched
        to the BOE line items by serial to populate the CTN column.

        Returns ``{ download_token, summary }`` on success, or the mapped
        ``{ error_code, message }`` body on rejection/failure.
        """
        # Validate the usd_rate form field up front (request validation -> 422).
        try:
            rate = float(usd_rate)
        except (TypeError, ValueError):
            return _error_response(
                ERROR_INVALID_USD_RATE, MESSAGE_INVALID_USD_RATE
            )

        raw = await file.read()
        filename = file.filename or "upload.pdf"

        invoice_raw: bytes | None = None
        if invoice is not None and invoice.filename:
            invoice_raw = await invoice.read()

        result = orchestrator.convert(raw, filename, rate, invoice_raw=invoice_raw)

        if not result.ok:
            return _error_response(
                result.error_code or ERROR_CONVERSION_FAILED,
                result.message or "The conversion could not be completed.",
            )

        return {
            "download_token": result.download_token,
            "summary": _summary_to_dict(result.summary),
            "output_complete": result.output_complete,
        }

    @app.get("/api/download/{download_token}")
    def download(download_token: str):
        """Stream the generated ``.xlsx`` bytes for a token, or 404 if unknown.

        NOTE (Milestone 1): unauthenticated endpoint -- local tool only.
        """
        workbook_bytes = orchestrator.get_download(download_token)
        if workbook_bytes is None:
            # Unknown/expired token => 404 per the Web_Interface contract.
            return JSONResponse(
                status_code=404,
                content={
                    "error_code": "NOT_FOUND",
                    "message": (
                        "The requested download was not found or expired."
                    ),
                },
            )

        return Response(
            content=workbook_bytes,
            media_type=XLSX_MEDIA_TYPE,
            headers={
                "Content-Disposition": (
                    'attachment; filename="bill_of_entry.xlsx"'
                )
            },
        )

    return app


# The single-file upload page. Plain HTML/JS, no build step: an upload control,
# a processing indicator shown while the request is in flight, and a results
# view that renders the summary and enables the download link (Req 1.1/1.4/1.8).
_UPLOAD_PAGE_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Bill of Entry Converter</title>
  <style>
    body { font-family: system-ui, sans-serif; max-width: 760px; margin: 2rem auto; padding: 0 1rem; color: #1a1a1a; }
    h1 { font-size: 1.4rem; }
    form { border: 1px solid #ddd; border-radius: 8px; padding: 1.25rem; }
    label { display: block; margin: 0.5rem 0 0.25rem; font-weight: 600; }
    input[type=number] { padding: 0.4rem; width: 12rem; }
    button { margin-top: 1rem; padding: 0.6rem 1.2rem; font-size: 1rem; cursor: pointer; }
    .hidden { display: none; }
    #processing { margin-top: 1rem; color: #555; }
    .spinner { display: inline-block; width: 1rem; height: 1rem; border: 2px solid #ccc;
      border-top-color: #333; border-radius: 50%; animation: spin 0.8s linear infinite; vertical-align: middle; }
    @keyframes spin { to { transform: rotate(360deg); } }
    #results { margin-top: 1.5rem; border-top: 1px solid #eee; padding-top: 1rem; }
    .error { color: #b00020; font-weight: 600; }
    .warn { color: #9a6700; }
    table { border-collapse: collapse; margin-top: 0.5rem; }
    td, th { border: 1px solid #ddd; padding: 0.3rem 0.6rem; text-align: left; }
    .download-link { display: inline-block; margin-top: 1rem; padding: 0.5rem 1rem;
      background: #0b5; color: #fff; text-decoration: none; border-radius: 6px; }
  </style>
</head>
<body>
  <h1>Bill of Entry &rarr; CTN Excel Converter</h1>
  <p>Upload a single Bill of Entry PDF and provide the USD rate. The converter
     produces the CTN-layout workbook for download.</p>

  <form id="convert-form">
    <label for="file">Bill of Entry PDF</label>
    <input type="file" id="file" name="file" accept="application/pdf" required />

    <label for="usd_rate">USD Rate (INR per USD)</label>
    <input type="number" id="usd_rate" name="usd_rate" step="any" min="0" required />

    <label for="invoice">Invoice / Packing List PDF (optional)</label>
    <input type="file" id="invoice" name="invoice" accept="application/pdf" />

    <div>
      <button type="submit" id="submit-btn">Convert</button>
    </div>
  </form>

  <div id="processing" class="hidden">
    <span class="spinner"></span> Converting&hellip; this may take up to a minute.
  </div>

  <div id="results" class="hidden"></div>

  <script>
    const form = document.getElementById('convert-form');
    const processing = document.getElementById('processing');
    const results = document.getElementById('results');
    const submitBtn = document.getElementById('submit-btn');

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      results.classList.add('hidden');
      results.innerHTML = '';
      processing.classList.remove('hidden');
      submitBtn.disabled = true;

      const fd = new FormData();
      fd.append('file', document.getElementById('file').files[0]);
      fd.append('usd_rate', document.getElementById('usd_rate').value);
      const invoiceFile = document.getElementById('invoice').files[0];
      if (invoiceFile) { fd.append('invoice', invoiceFile); }

      try {
        const resp = await fetch('/api/convert', { method: 'POST', body: fd });
        const data = await resp.json();
        processing.classList.add('hidden');
        submitBtn.disabled = false;

        if (!resp.ok) {
          renderError(data.message || 'The conversion could not be completed.');
          return;
        }
        renderSummary(data);
      } catch (err) {
        processing.classList.add('hidden');
        submitBtn.disabled = false;
        renderError('Network or server error: ' + err);
      }
    });

    function renderError(message) {
      results.classList.remove('hidden');
      results.innerHTML = '<p class="error">' + escapeHtml(message) + '</p>';
    }

    function renderSummary(data) {
      const s = data.summary || {};
      results.classList.remove('hidden');
      let html = '<h2>Conversion complete</h2>';
      if (data.output_complete === false) {
        html += '<p class="warn">Note: the output may be incomplete &mdash; ' +
                'review the discrepancies below.</p>';
      }
      html += '<table>' +
        row('Line items extracted', s.line_items_extracted) +
        row('Declared item count', s.declared_item_count) +
        row('Total invoice amount (USD)', s.total_invoice_amount_usd) +
        row('Declared invoice amount (USD)', s.declared_invoice_amount_usd) +
        row('Fields flagged for review', s.review_flag_count) +
        '</table>';

      if (s.discrepancies && s.discrepancies.length) {
        html += '<h3>Discrepancies</h3><ul>';
        for (const d of s.discrepancies) {
          html += '<li class="warn">[' + escapeHtml(d.kind) + '] ' +
                  escapeHtml(d.message) + '</li>';
        }
        html += '</ul>';
      }

      const url = '/api/download/' + encodeURIComponent(data.download_token);
      html += '<a class="download-link" href="' + url + '">Download Excel workbook</a>';
      results.innerHTML = html;
    }

    function row(label, value) {
      const v = (value === null || value === undefined) ? '' : value;
      return '<tr><th>' + escapeHtml(label) + '</th><td>' + escapeHtml(String(v)) + '</td></tr>';
    }

    function escapeHtml(str) {
      return String(str).replace(/[&<>"']/g, (c) => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
      }[c]));
    }
  </script>
</body>
</html>
"""
