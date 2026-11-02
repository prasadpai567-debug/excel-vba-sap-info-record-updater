# Excel VBA SAP Purchasing Info Record Updater

An Excel VBA automation example that updates SAP purchasing info records through SAP GUI Scripting and transaction ME12.

The project is aimed at procurement operations where supplier, material, purchasing-organization, plant, and price values must be entered repeatedly in Excel and SAP.

> Portfolio scope: this public repository contains exported VBA source and synthetic input data. The original macro-enabled workbook is intentionally not included because it may contain company, system, or business data.

## Business problem

Maintaining purchasing info records manually requires repeated data entry, SAP screen navigation, and manual copying of results back into a tracking sheet. This creates avoidable effort and makes it harder to identify which records succeeded or failed during a batch.

## How the automation helps

The macro reads input rows from Excel, validates required fields, opens SAP transaction ME12, enters the purchasing-info-record key, updates the price, and writes the returned info-record number and SAP status back to the worksheet.

The benefits described here are qualitative. This repository contains no measured time-savings or error-rate results.

## Implemented features

The following behavior is verified in Module1.bas:

- Manual execution through the public procedure SAP_Final_Solution_V2.
- Processing from row 2 through the last non-empty supplier-number row.
- Validation of supplier, material, purchasing organization, plant, and numeric price.
- Price formatting to two decimal places before SAP entry.
- Connection to the first active SAP GUI connection and first SAP session.
- SAP GUI Scripting navigation to transaction ME12.
- Info-record number written to column F.
- SAP status or processing result written to column H.
- Optional popup handling during navigation and save confirmation.
- Row-level error handling that marks a failed row and continues with later rows.
- Color-coded processing, success, and failure states.
- Completion and connection-error message boxes.

The source does not implement authentication, SAP-session selection, scheduling, external logging, rollback, a UserForm, or a dry-run mode.

## Repository contents

| File | Purpose |
|---|---|
| Module1.bas | Main SAP automation procedure and helper procedures |
| Sheet1.cls | Exported worksheet document metadata; no event logic |
| ThisWorkbook.cls | Exported workbook document metadata; no event logic |
| sample_input.csv | Synthetic input example |
| .gitignore | Excludes macro-enabled workbooks, OLE binaries, temporary files, and logs |
| README.md | Project, setup, workflow, limitation, and portfolio documentation |

The repository does not include an executable XLSM workbook, worksheet shape, screenshots, automated tests, or a GitHub Actions workflow.

## Workbook layout

| Column | Header | Use |
|---|---|---|
| A | Supplier No. | SAP supplier number; also determines the last row |
| B | Description | Informational only; not sent to SAP |
| C | Material Number | SAP material number |
| D | Purchasing Org. | SAP purchasing organization |
| E | Plant | SAP plant |
| F | inforecord number | Written by the macro after SAP returns a value |
| G | Price | Numeric price sent to SAP |
| H | SAP Status | Processing result or SAP status-bar text |

sample_input.csv is input-only, so columns F and H are generated in the Excel workbook during processing.

## Application workflow

1. Use the active Excel worksheet.
2. Connect to the first available SAP GUI connection and session.
3. Find the last used row from column A.
4. Set each row to Processing... and validate the required values.
5. Open ME12 and enter supplier, material, purchasing organization, and plant.
6. Read the SAP info-record number into column F.
7. Navigate to the purchasing-organization data and enter the price.
8. Validate and save the SAP record.
9. Write the SAP status to column H and continue with the next row.

## Technologies and requirements

- Microsoft Excel VBA and macro-enabled workbooks.
- SAP GUI Scripting for Windows.
- Late-bound COM automation through VBA Object variables.
- SAP transaction ME12.
- Windows desktop Excel with VBA support.
- An active logged-in SAP GUI session.
- SAP GUI Scripting enabled on the client and SAP server.
- Authorization to access ME12 and change the relevant purchasing info records.
- A workbook and SAP screen layout compatible with the hard-coded control IDs.

## Setup from the source-only repository

1. Clone or download this repository.
2. Create a blank macro-enabled Excel workbook in a controlled location.
3. Open the VBA editor and import Module1.bas into a standard VBA module.
4. Create the worksheet columns shown above.
5. Add a worksheet shape or button and assign it to SAP_Final_Solution_V2.
6. Use synthetic data or an approved non-production SAP environment.
7. Open SAP GUI and leave the intended session active.
8. Confirm scripting and authorization before testing.
9. Run one test row first, then review columns F and H.

Sheet1.cls and ThisWorkbook.cls contain only exported document-module attributes and are not required for the current macro behavior. Do not publish the original business workbook or its embedded OLE/VBA project without sanitizing it.

Do not lower Excel macro-security settings globally. Use an approved trusted location or per-file process.

## Usage example

| Supplier No. | Description | Material Number | Purchasing Org. | Plant | Price |
|---:|---|---|---:|---:|---:|
| 100000 | Demo Supplier | DEMO-MAT-001 | 1000 | 1000 | 12.50 |

The connected SAP system, master data, authorizations, and screen configuration determine the actual response.

## Expected output

| Column | Example result | Meaning |
|---|---|---|
| F | <SAP-generated-info-record> | Info-record number read from SAP |
| H | <SAP status message> | SAP status-bar text or processing result |

For an invalid row, column H may contain a message such as Skipped: Price is blank or invalid. These examples describe the output shape and do not claim a successful SAP run.

## Error handling and logging

- Missing SAP connection or session: displays a connection-error message and stops the run.
- Invalid required input: writes a row-level message to column H, fills the row red, and continues.
- SAP error during a valid row: writes the VBA error number and description to column H, marks the row red, and continues.
- SAP busy state: waits with Session.Busy, DoEvents, and a 30-second timeout.
- No separate log file, run ID, timestamp, operator record, or external audit history is generated.
- The macro does not roll back an SAP change if a later step in the same row fails.

### Known implementation risk

The current status helper functions use permissive error handling when reading the SAP status bar. If the status bar cannot be read, an empty status can be replaced with Updated successfully. Treat the current output as operator feedback, not as an independent audit record, until this behavior is hardened and validated.

## Security and data handling

- No credentials, passwords, tokens, or connection strings are present in the tracked VBA source.
- The sample CSV uses demonstration-looking values.
- Keep supplier, material, plant, price, SAP system, user, screenshot, and log data private unless approved and sanitized.
- Test with synthetic data or an approved non-production SAP environment before changing live records.
- The macro writes to SAP; confirm the active workbook, SAP connection, SAP session, and target data before running it.

## Current limitations

- No runnable XLSM workbook is included.
- The macro uses the active worksheet and the first SAP connection/session.
- SAP control IDs are hard-coded.
- The last row is determined only from column A.
- Currency, price unit, validity dates, condition type, and other purchasing fields are not handled.
- There is no session selector, dry-run mode, retry queue, external audit log, or rollback.
- No automated tests or measured business metrics are included.

## Future improvements

These items are not implemented:

- Add preflight checks and explicit SAP-session selection.
- Replace permissive status handling with explicit success/failure confirmation.
- Move control IDs and mappings into documented configuration.
- Add stronger validation for Excel error values, price rules, duplicates, and blank-row gaps.
- Add recoverable timeout handling, reconciliation, and an append-only run log where policy permits.
- Add mock-based tests for validation and output-writing logic.
- Add sanitized screenshots and a demonstration workbook only after review.

## Screenshots and portfolio summary

No screenshots are currently included. Suggested sanitized additions are an input worksheet, RUN button, processing status, completed output, and redacted ME12 screen. Never include supplier data, company names, SAP identifiers, usernames, client numbers, live screenshots, or exported logs without approval and sanitization.

> Developed an Excel VBA automation that batch-updates SAP purchasing info records through SAP GUI Scripting. Implemented structured input validation, SAP transaction navigation, price updates, status capture, row-level error handling, and color-coded results to reduce repetitive procurement-maintenance work and improve processing visibility.

## License and disclaimer

No open-source license is currently included. Select a license that matches the ownership and publication rights for the VBA source before allowing reuse.

Test this automation with synthetic data or an approved non-production SAP environment before using it against live records.
