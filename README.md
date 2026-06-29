# Bill of Entry Converter (Milestone 1)

Convert an Indian Customs **Bill of Entry (BOE)** PDF into the pre-defined CTN
Excel workbook layout (`1357 ctn llp.xlsx`). Milestone 1 focuses on faithful
parsing: every header field and line item is captured verbatim, derived
monetary values are computed at full precision, and any anomaly is surfaced to
the user rather than silently dropped.

## Project layout

```
boe_converter/
  models.py        # shared frozen dataclasses (the contracts between components)
  parser.py        # PDF_Parser: orientation-aware extraction (tasks 4.x / 5.x)
  calculator.py    # Value_Calculator: pure per-line + totals math (tasks 2.x)
  excel_writer.py  # Excel_Generator: exact CTN layout (tasks 7.x)
  validator.py     # UploadValidator: size / PDF / BOE checks (task 9.1)
  orchestrator.py  # ConversionOrchestrator: sequence + cross-checks (task 9.3)
  web/             # FastAPI Web_Interface (task 11.1)
tests/             # pytest + hypothesis tests
```

## Setup

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
```

## Running tests

```bash
pytest
# thorough property-test run
pytest --hypothesis-profile=ci
```

The default Hypothesis profile runs `max_examples=100` per property.
