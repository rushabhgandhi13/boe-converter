# Requirements Document

## Introduction

This feature is a web-based tool that converts an Indian Customs **Bill of Entry (BOE)** PDF into a
specific, pre-defined Excel workbook (the "CTN" format) used by the importer's accounting team. Today
this conversion is fully manual: a staff member reads the BOE PDF, copies each line item and header
field into a spreadsheet, simplifies item names, and computes landed-cost figures (customs duty, IGST,
surcharge, land cost with/without GST) by hand. The tool automates that conversion.

The scope of this document is **Milestone 1** only: read a BOE PDF accurately and produce the target
Excel in the exact layout of the supplied sample (`1357 ctn llp.xlsx`). Parsing accuracy is the primary,
non-negotiable requirement — every field, including anomalies and outliers, must be captured faithfully,
and no line item may be silently dropped or altered.

**Out of scope (future Milestone B):** importing/exporting the resulting data into the Tally accounting
application (whether through Tally APIs or a Tally-importable Excel format). Tally integration is noted
here only to confirm it is deliberately excluded from Milestone 1.

The reference inputs/outputs used to derive these requirements are:
- Source BOE PDF: `205090022062026INNSA1BE0230620261842.pdf` (30 pages, 45 line items)
- Target Excel: `1357 ctn llp.xlsx` (Sheet1, header block + 45-row line-item table + totals + auxiliary sections)

## Glossary

- **BOE (Bill of Entry)**: A customs clearance document filed for imported goods. Source input for this tool.
- **Converter**: The overall system being built that transforms a BOE PDF into the target Excel workbook.
- **PDF_Parser**: The component that reads a BOE PDF and extracts structured header and line-item data.
- **Excel_Generator**: The component that writes extracted and computed data into the target Excel layout.
- **Value_Calculator**: The component that computes derived numeric fields (duties, IGST, land cost, etc.).
- **Web_Interface**: The browser-based UI through which a User uploads a BOE PDF and downloads the Excel.
- **User**: The importer's staff member who operates the Converter.
- **Line_Item**: One imported product entry in the BOE (identified by its Item Serial Number).
- **Header_Block**: The top section of the target Excel (rows 1-8) containing document-level fields
  (Company name, Party Name, USD Rate, Details, Invoice No/Date, BE No/Date, B/L No/Date, etc.).
- **Item_Table**: The line-item table in the target Excel beginning at the header row (row 12).
- **Totals_Row**: The summary row in the target Excel that aggregates numeric columns (row 61 in the sample).
- **CTH / HSN Code**: Customs Tariff Heading / Harmonized System Nomenclature classification code for an item.
- **Assessable_Value**: The customs-assessed value (INR) of a Line_Item, shown as "ASSESS VALUE" in the BOE.
- **BCD (Basic Customs Duty)**: A customs duty levied as a percentage of the Assessable_Value.
- **SWS (Social Welfare Surcharge)**: A surcharge levied at 10% of the BCD amount.
- **IGST (Integrated GST)**: Goods and Services Tax charged on (Assessable_Value + customs duties).
- **USD_Rate**: The exchange rate (INR per USD) applied to convert invoice USD values to INR.
- **Unit_Price_USD**: The per-unit invoice price in USD (the BOE "UPI - Unit Price Invoiced" field).
- **Land_Cost**: The computed landed cost of an item, expressed both excluding and including GST.
- **Tally**: The external accounting application targeted by the out-of-scope future Milestone B.

## Requirements

### Requirement 1: Upload and convert a BOE PDF through a web interface

**User Story:** As a User, I want to upload a BOE PDF in a web browser and receive the converted Excel file, so that I no longer perform the conversion manually.

#### Acceptance Criteria

1. THE Web_Interface SHALL provide a control that accepts exactly one PDF file per submission as input.
2. IF the uploaded file exceeds 50 MB, THEN THE Converter SHALL reject the file, display a message stating that the file exceeds the 50 MB size limit, and SHALL NOT produce an output file.
3. WHEN a User submits a valid BOE PDF, where a valid BOE PDF is a readable PDF whose extracted content is recognized as a Bill of Entry containing the fields required to populate the sample CTN workbook layout, THE Converter SHALL produce a downloadable Excel workbook in the target CTN layout within 60 seconds of submission.
4. WHILE a conversion is in progress, THE Web_Interface SHALL display a processing indicator to the User, and SHALL continue displaying it until the conversion completes or an error is reported.
5. IF the uploaded file is not a readable PDF, THEN THE Converter SHALL reject the file, display a message stating that a valid PDF is required, and SHALL NOT produce an output file.
6. IF the uploaded PDF is not recognizable as a Bill of Entry, THEN THE Converter SHALL display a message stating that the document is not a recognized Bill of Entry and SHALL NOT produce an output file.
7. IF conversion fails after the uploaded PDF is recognized as a Bill of Entry, THEN THE Converter SHALL display a message indicating that the conversion could not be completed and SHALL NOT make any partial or incomplete Excel workbook available for download.
8. WHEN conversion completes, THE Web_Interface SHALL allow the User to download the generated Excel workbook.

### Requirement 2: Accurately extract all Bill of Entry header fields

**User Story:** As a User, I want every document-level field read correctly from the BOE, so that the Excel header block matches the source document.

#### Acceptance Criteria

1. THE PDF_Parser SHALL extract the Bill of Entry Number from the BOE.
2. THE PDF_Parser SHALL extract the Bill of Entry Date from the BOE.
3. THE PDF_Parser SHALL extract the Invoice Number from the BOE.
4. THE PDF_Parser SHALL extract the Invoice Date from the BOE.
5. THE PDF_Parser SHALL extract the total Invoice Amount and its currency from the BOE.
6. THE PDF_Parser SHALL extract the total number of packages (PKG / CTN count) from the BOE.
7. THE PDF_Parser SHALL extract the supplier/exporter name (Party Name) from the BOE.
8. THE PDF_Parser SHALL extract the container details (container number and count) from the BOE.
9. WHERE a Bill of Lading number and date are present in the BOE, THE PDF_Parser SHALL extract the Bill of Lading Number and Bill of Lading Date.
10. IF a header field required by criteria 1 through 8 cannot be located in the BOE, or its printed characters cannot be resolved into a value, THEN THE Converter SHALL record that field as missing, SHALL report each such missing field to the User identified by its field name, and SHALL NOT substitute an inferred or default value.
11. THE PDF_Parser SHALL record each extracted header field as the value printed in the BOE, preserving its characters, digits, and currency symbol without reformatting, truncating, rounding, or inferring content.

### Requirement 3: Accurately extract all Bill of Entry line items

**User Story:** As a User, I want every line item read correctly and completely, so that no product, quantity, or value is dropped or altered.

#### Acceptance Criteria

1. THE PDF_Parser SHALL extract one record for each Line_Item present in the BOE.
2. WHEN the BOE declares a total item count, THE PDF_Parser SHALL extract a number of Line_Item records equal to that declared count.
3. IF the number of extracted Line_Item records does not equal the BOE's declared item count, THEN THE Converter SHALL report the discrepancy to the User, stating both the BOE's declared item count and the number of records extracted, and SHALL NOT present the output as complete.
4. FOR each Line_Item, THE PDF_Parser SHALL extract the Item Serial Number.
5. FOR each Line_Item, THE PDF_Parser SHALL extract the CTH / HSN Code.
6. FOR each Line_Item, THE PDF_Parser SHALL extract the complete item description text verbatim, without truncation, abbreviation, or omission.
7. FOR each Line_Item, THE PDF_Parser SHALL extract the Unit_Price_USD.
8. FOR each Line_Item, THE PDF_Parser SHALL extract the declared quantity and the unit of quantity.
9. FOR each Line_Item, THE PDF_Parser SHALL extract the Assessable_Value.
10. FOR each Line_Item, THE PDF_Parser SHALL extract the BCD rate and BCD amount.
11. FOR each Line_Item, THE PDF_Parser SHALL extract the IGST rate.
12. FOR each Line_Item, THE PDF_Parser SHALL extract the total duty amount.
13. WHERE a Line_Item spans more than one PDF page, THE PDF_Parser SHALL merge the Line_Item's fields from all spanned pages into a single record with each field value appearing exactly once, without duplication or omission.
14. WHERE a Line_Item has an exemption notification that produces a zero duty rate, THE PDF_Parser SHALL extract the corresponding rate and amount as the numeric value zero rather than leaving them blank.
15. IF a Line_Item field required by criteria 4 through 12 cannot be located in the BOE, or its printed characters cannot be resolved into a value, THEN THE Converter SHALL record that field as missing, SHALL report it to the User identified by its Line_Item and field name, SHALL NOT substitute an inferred or default value, and SHALL NOT drop the Line_Item.

### Requirement 4: Populate the Excel header block from extracted data

**User Story:** As a User, I want the Excel header block filled from the BOE, so that the output identifies the shipment correctly.

#### Acceptance Criteria

1. THE Excel_Generator SHALL write the extracted supplier name into the Party Name field of the Header_Block.
2. THE Excel_Generator SHALL write the extracted Invoice Number and Invoice Date into the Header_Block exactly as extracted, without reformatting the number or date values.
3. THE Excel_Generator SHALL write the extracted Bill of Entry Number and Bill of Entry Date into the Header_Block exactly as extracted, without reformatting the number or date values.
4. WHERE a Bill of Lading Number and Bill of Lading Date are present in the extracted data, THE Excel_Generator SHALL write them into the Header_Block exactly as extracted, without reformatting the number or date values.
5. THE Excel_Generator SHALL write the USD_Rate into the Header_Block exactly as extracted, without altering its numeric precision.
6. THE Excel_Generator SHALL write the package/container details (e.g., "CO-32 CTN-1357") into the Details field of the Header_Block.
7. THE Excel_Generator SHALL write the importer company name into the Company name field of the Header_Block.
8. WHERE a Header_Block field has no corresponding value available from the BOE or configuration (for example Eway bill number/date and Remittance date/rate), THE Excel_Generator SHALL leave that field's cell empty with no characters, spaces, or placeholder text.
9. IF a source value was flagged as missing or unreadable during extraction, THEN THE Excel_Generator SHALL leave the corresponding cell empty with no characters, spaces, or placeholder text, rather than writing an inferred or placeholder value.

### Requirement 5: Populate the Excel line-item table from extracted data

**User Story:** As a User, I want each BOE line item written to the correct columns of the Item_Table, so that the output matches the sample layout.

#### Acceptance Criteria

1. THE Excel_Generator SHALL write Line_Item records into the Item_Table in ascending Item Serial Number order, beginning at the first data row immediately below the Item_Table header row (row 12), writing exactly one Line_Item record per row with no blank rows between records.
2. THE Excel_Generator SHALL write a sequential Sr. no. that equals 1 in the first Line_Item row and increases by exactly 1 in each subsequent Line_Item row, such that the Sr. no. values form the consecutive sequence 1 through N with no gaps, repeats, or skipped values for N Line_Items.
3. THE Excel_Generator SHALL write the extracted item description into the Description column.
4. THE Excel_Generator SHALL write the extracted CTH / HSN Code into the HSN CODE column.
5. THE Excel_Generator SHALL write the extracted Unit_Price_USD into the Unit Price in USD column.
6. THE Excel_Generator SHALL write the extracted quantity into the QTY column.
7. THE Excel_Generator SHALL write the extracted unit into the Unit column.
8. WHERE a Line_Item's unit, after trimming leading and trailing whitespace and ignoring letter case, equals "DOZ", THE Excel_Generator SHALL write the piece count (QTY multiplied by 12) into the pcs column.
9. WHERE a Line_Item's unit, after trimming leading and trailing whitespace and ignoring letter case, does not equal "DOZ", THE Excel_Generator SHALL leave the pcs column blank.
10. THE Excel_Generator SHALL write the extracted Assessable_Value into the CUSTOM ASS VALUE column.
11. IF a value required for any Item_Table cell of a Line_Item is missing or unreadable, THEN THE Excel_Generator SHALL leave only that cell blank, SHALL flag that cell as requiring User review, and SHALL write all remaining cells of that Line_Item's row without omitting the row.
12. THE Excel_Generator SHALL reproduce the Item_Table column order and each header label as a character-for-character exact match of the sample Item_Table, including spacing, punctuation, and letter case.

### Requirement 6: Compute derived per-line monetary values

**User Story:** As a User, I want duties, taxes, and landed costs calculated automatically per line, so that I do not compute them by hand.

#### Acceptance Criteria

1. THE Value_Calculator SHALL compute the per-line invoice Amount (USD) as Unit_Price_USD multiplied by QTY.
2. THE Value_Calculator SHALL compute the per-line purchase value in INR as the invoice Amount (USD) multiplied by the USD_Rate.
3. THE Value_Calculator SHALL compute the per-line SWS amount as the BCD amount multiplied by 0.10.
4. THE Value_Calculator SHALL compute the per-line total customs duty (excluding IGST) as the sum of the BCD amount and the SWS amount.
5. THE Value_Calculator SHALL compute the per-line IGST amount as the IGST rate, expressed as a decimal fraction (for example 0.18 for an 18% rate), multiplied by the sum of the Assessable_Value and the total customs duty (excluding IGST).
6. THE Value_Calculator SHALL compute the per-line combined duty as the sum of the total customs duty (excluding IGST) and the IGST amount.
7. THE Value_Calculator SHALL compute the per-line Land_Cost excluding GST as the sum of the purchase value in INR and the total customs duty (excluding IGST).
8. THE Value_Calculator SHALL compute the per-line Land_Cost including GST as the sum of the Land_Cost excluding GST and the IGST amount.
9. THE Value_Calculator SHALL compute the per-line purchase rate per unit as the Land_Cost excluding GST divided by QTY.
10. IF a Line_Item's QTY is zero, THEN THE Value_Calculator SHALL set the per-line purchase rate per unit to zero rather than performing a division by zero.
11. THE Excel_Generator SHALL write each computed value into its corresponding Item_Table column (Amount, Rate Per USD in purchase, SURCHARGE, TOTAL Custom Duty, GST, total custom duty, LAND COST OF PURCHASE WITHOUT GST, LAND COST OF PURCHASE WITH GST, Rate purchase per unit).
12. THE Value_Calculator SHALL retain every per-line computed value at full floating-point precision without rounding.
13. IF any input required to compute a per-line derived value (Unit_Price_USD, QTY, USD_Rate, BCD amount, IGST rate, or Assessable_Value) is missing or non-numeric, THEN THE Value_Calculator SHALL leave the dependent computed value blank and SHALL flag the affected Line_Item as requiring User review rather than substituting a default value.

### Requirement 7: Compute and write the totals row

**User Story:** As a User, I want a totals row that sums the columns, so that I can verify the workbook against the BOE invoice total.

#### Acceptance Criteria

1. THE Value_Calculator SHALL compute the total invoice Amount (USD) as the sum of all per-line invoice Amounts (USD).
2. THE Value_Calculator SHALL compute the column totals for Assessable_Value, total customs duty, IGST, and both Land_Cost columns as the sum of the corresponding per-line values.
3. WHEN no Line_Item records are present, THE Value_Calculator SHALL set every computed column total in the Totals_Row to zero.
4. THE Excel_Generator SHALL write the computed totals into the Totals_Row in the columns matching the sample layout, without altering their numeric precision.
5. THE Excel_Generator SHALL write the total package/CTN count, taken from the extracted BOE total number of packages, into the Totals_Row.
6. WHEN the Totals_Row is generated, IF the computed total invoice Amount (USD) and the BOE's declared total Invoice Amount differ by more than 0.01 USD, THEN THE Converter SHALL display a message to the User indicating the discrepancy and showing both the computed total and the declared total, and SHALL retain the generated workbook.
7. WHEN the Totals_Row is generated, IF the BOE's declared total Invoice Amount is absent or unreadable, THEN THE Converter SHALL display a message to the User indicating that the computed invoice total could not be verified against the BOE.

### Requirement 8: Preserve numeric fidelity and reproduce the exact output layout

**User Story:** As a User, I want the output workbook to match the sample format precisely with no rounding errors, so that downstream use and review are reliable.

#### Acceptance Criteria

1. THE Excel_Generator SHALL produce exactly one worksheet, named "Sheet1", and SHALL NOT add any additional worksheets to the output workbook.
2. THE Excel_Generator SHALL write the Header_Block at rows 1 through 8, the Item_Table header row at row 12, the line items in the rows immediately following the Item_Table header row, and the Totals_Row at the same row position as the sample workbook (row 61 in the sample), using the identical cell column positions of the sample workbook.
3. THE Excel_Generator SHALL reproduce the auxiliary section label text exactly as it appears in the sample ("DETAILS AS PER CHALLANS", "DETAILS AS PER TALLY", "CLEARANCE AND FORWARDING INVOICE", and the C&F detail block) at the same cell positions as the sample, and SHALL leave the data cells within those sections empty (containing no value).
4. THE Value_Calculator SHALL retain each value extracted directly from the BOE as its exact extracted numeric value, without applying rounding, truncation, or reformatting before it is written to the output.
5. THE Excel_Generator SHALL write each computed value to the output at the full numeric precision produced by the Value_Calculator, without rounding or truncation.
6. WHERE the sample workbook leaves a column empty for all line items (for example PARTY NAME, BILLING AMOUNT), THE Excel_Generator SHALL leave that column empty (containing no value and no whitespace) for all line items in the output.

### Requirement 9: Report parsing anomalies without dropping data

**User Story:** As a User, I want to be told about anything unusual in the BOE instead of having it silently omitted, so that I can trust the output's completeness.

#### Acceptance Criteria

1. WHEN the Converter completes a conversion, THE Converter SHALL present a summary stating the number of Line_Items extracted, the total invoice Amount in USD, and the count of fields flagged as requiring User review.
2. IF a Line_Item field cannot be parsed into its expected data type, THEN THE Converter SHALL write the raw extracted text for that field into the output without removing the field.
3. IF a Line_Item field cannot be parsed into its expected data type, THEN THE Converter SHALL mark that field with a visible indication that it requires User review.
4. IF a numeric value extracted from the BOE differs from the same value recomputed from its related fields by more than 0.01 in that value's unit, THEN THE Converter SHALL report both the extracted value and the recomputed value to the User.
5. THE Converter SHALL NOT remove a Line_Item from the output solely because one of its fields failed to parse.

## Source-to-Target Field Mapping (informative)

This table records the field-by-field mapping derived from the sample PDF and Excel. It classifies each
target column as **Direct** (copied from the PDF), **Computed** (derived by the Value_Calculator), or
**Manual/External** (not available from the BOE alone — see Open Questions). The exact computed formulas
are stated normatively in Requirement 6.

| Target column (Item_Table) | Source classification | Notes |
|---|---|---|
| Sr. no. | Computed | Sequential 1..N |
| PARTY NAME | Manual/External | Blank for all line items in sample |
| BILLING AMOUNT | Manual/External | Blank for all line items in sample |
| AS PER TALLY NAME | Manual/External | Simplified item name; appears human-entered (see Open Q3) |
| Description | Direct | BOE item description (CTH item text) |
| HSN CODE | Direct | BOE CTH |
| CTN | Manual/External | Cartons per line; not clearly in BOE (see Open Q1) |
| QTY | Direct | BOE quantity (see Open Q2 on rounding/which qty field) |
| Unit | Direct | BOE unit (DOZ/KGS/NOS/PCS) |
| pcs | Computed | QTY×12 when Unit = DOZ, else blank |
| Unit Price in USD | Direct | BOE UPI |
| Amount | Computed | Unit Price × QTY |
| Rate Per USD in purchase | Computed | Amount × USD_Rate |
| CUSTOM ASS VALUE | Direct | BOE Assessable Value (INR) |
| LAND COST OF PURCHASE WITHOUT GST | Computed | Purchase INR + (BCD + SWS) |
| TOTAL Custom Duty | Computed | BCD + SWS |
| GST | Computed | IGST rate × (Assessable Value + BCD + SWS) |
| RATE OF DUTY IGST | Direct | BOE IGST rate |
| total custom duty | Computed | TOTAL Custom Duty + GST |
| RATE OF INTEREST (col T) | Direct | Actually the BCD rate (label is misleading) |
| CUST AIDC (col U) | Direct | Actually the BCD amount (label is misleading) |
| RATE OF INTEREST (col V) | Constant | SWS rate = 0.10 |
| SURCHARGE | Computed | BCD × 0.10 (SWS) |
| LAND COST OF PURCHASE WITH GST | Computed | Land cost without GST + GST |
| Rate purchase per unit | Computed | Land cost without GST ÷ QTY |

> Note: Several sample column headers are misleading (e.g., "CUST AIDC" holds the BCD amount; "RATE OF
> INTEREST" holds the BCD rate / SWS rate). The mapping above reflects the **actual** computed behavior
> observed in the sample, which the implementation must reproduce. This should be confirmed (Open Q4).

## Open Questions / Assumptions

These items were identified while reverse-engineering the sample and need User confirmation. Current
working assumptions are stated so the initial requirements remain actionable; answers may refine
Requirements 5-7.

1. **CTN (cartons per line item):** The per-line CTN values (e.g., 5, 45, 17) do not appear to come from
   the BOE line items; only the grand total (1357) matches the BOE package count. Are per-line CTN values
   taken from a separate packing list, entered manually, or derived by a rule? *Assumption:* manual/external for Milestone 1.
2. **QTY source and rounding:** The sample QTY appears to use a rounded quantity (e.g., 255 vs the BOE's
   255.15 KGS). Which BOE quantity field should drive QTY (commercial qty vs standard qty), and what
   rounding rule applies? *Assumption:* use the commercial quantity as printed, no rounding, pending confirmation.
3. **AS PER TALLY NAME:** This simplified name (with sample typos like "Decorative Itemsss") looks
   human-entered. Should the tool leave it blank for manual entry, or derive it from the description by a rule/lookup? *Assumption:* leave blank/manual for Milestone 1.
4. **Misleading column labels / duty model:** Confirm that "CUST AIDC" should hold the BCD amount and the
   two "RATE OF INTEREST" columns hold BCD rate and SWS rate (10%) respectively, and that AIDC/interest are
   not separately required. Confirm SWS is always 10% of BCD.
5. **USD_Rate origin:** Is the USD_Rate (95.3 in the sample) read from the BOE, supplied by the User per
   conversion, or read from a configuration value? *Assumption:* User-supplied per conversion.
6. **Company name:** Is the importer company name ("Gemini Unicom LLP") a fixed configuration value or
   read from the BOE? *Assumption:* configuration value.
7. **Rounding/precision of computed values:** Should computed monetary values be stored at full floating
   precision (as in the sample) or rounded to a fixed number of decimals? *Assumption:* full precision, matching the sample.
8. **Multi-invoice / multi-page BOEs:** The sample BOE has a single invoice. Should Milestone 1 support a
   BOE containing multiple invoices, or is single-invoice sufficient? *Assumption:* single-invoice sufficient for Milestone 1.
9. **Auxiliary sections (Challans, Tally, C&F):** Confirm these lower sections should be emitted as empty
   templates in Milestone 1 (data entry/Tally linkage deferred to Milestone B). *Assumption:* empty templates.
