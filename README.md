# DTSC Report Wizard

This Google Sheets Apps Script helps generate DTSC report files and import invoice tab data automatically.

It also supports a separate month-total compile workflow for ALL Items Report (internal key `ALL_ITEMS`), DTSC, and NETCOST.

## What this tool does

1. Finds the customer folder in Purchases.
2. Creates/uses a Generated Reports folder.
3. Copies report templates:
   - ALL Items Report (internal key `ALL_ITEMS`)
   - DTSC
   - NETCOST (only when SB-20 is checked)
4. Finds invoice files by invoice folder name.
5. Ensures invoice has a PO tab.
6. Imports/refreshes:
   - EWASTE-PO
   - SB20-PO (only when SB-20 is checked)
7. Freezes those tabs to values.
8. Sends tabs into generated report files using this mapping:
   - PO -> ALL Items Report (internal key `ALL_ITEMS`)
   - EWASTE-PO -> DTSC
   - SB20-PO -> NETCOST
9. Updates IMPORTED-DATA in each report:
   - Writes imported tab names in A2:A31
   - Keeps multiple invoice rows (appends up to 30)

## Run the report wizard

1. Open the Google Sheet with this script attached.
2. Refresh the sheet.
3. Go to DTSC Report Wizard -> Start Report Wizard.
4. In the sidebar, keep Action as Run Wizard.
5. Enter:
   - Customer Number
   - Invoice Number(s) (comma or new line)
   - Invoice date format in invoice number should be `...-MM-DD-YY` or `...-MM-DD-YYYY` for correct month naming (example: `#232-01-02-26` => January 2026)
   - SB-20 checkbox
   - Mode (TEST or LIVE)
6. Click Start Run.

## Compile Month Totals (new)

Month total posting is now a separate manual action.

Use this when you want to fill month rows from existing monthly report files.

1. Open DTSC Report Wizard -> Start Report Wizard.
2. In the sidebar, set Action to Compile Month Totals.
3. Enter/select:
   - Customer Number
   - Mode (TEST or LIVE)
   - Report Type (ALL Items Report / `ALL_ITEMS`, DTSC, NETCOST)
   - Report Year
4. Click Load Report Files.
5. Pick the target report file.
6. Check one or more months to compile.
7. Click Start Compile and watch realtime progress/logs.

### How Compile Month Totals works

- It reads each selected month from that month’s source report file for the selected report type.
- It takes totals from row 20 in the source report and writes them into the matching month row in the target report.
- Existing values in selected month rows are overwritten.
- If a source month report is missing, that month is logged as missing and skipped.
- TEST and LIVE are handled separately based on selected mode.

### Monthly ranges used

- NETCOST: source `B20:R20` -> month row in `B25:R36` (month labels in `A25:A36`)
 - ALL Items Report (`ALL ITEMS`): source `B20:N20` -> month row in `B25:N36` (month labels in `A25:A36`)
- DTSC (`DTSC`): source `B20:L20` -> month row in `B25:L36` (month labels in `A25:A36`)

## TEST vs LIVE

- TEST mode:
  - Creates invoice test copies in Generated Reports - TEST
  - Original invoice files stay unchanged
- LIVE mode:
  - Updates original invoice files

## Required access

Each user must have access to:

- Purchases folder
- Customer folders and invoice files
- Report template files
- EWASTE-PO and SB20-PO source spreadsheet files

If access is missing, that step will fail for that user.

## Troubleshooting

### Exceeded maximum execution time

- Run fewer invoices per batch.
- Use exact invoice folder names.
- Confirm invoice folders exist under the selected customer tree.

### EWASTE-PO/SB20-PO not visible in invoice

- In TEST mode, changes are made to test copy files, not originals.
- Open the file shown in progress logs.

### Missing imports in reports

- Confirm invoice contains required PO tab.
- Confirm SB-20 checkbox is correct for that customer.

### Compile Month Totals did not update a month

- Confirm a source report exists for that month, year, report type, and mode.
- Confirm you selected the correct target report file.
- Check the progress log for missing-source month messages.

## Copy/paste summary for Google Sheet tab

Use this short version in a Sheet note/tab:

- Open DTSC Report Wizard -> Start Report Wizard
- Action = Run Wizard to generate/import reports
- Action = Compile Month Totals to post month rows
- Select TEST (safe) or LIVE (real updates)
- Watch progress/logs in realtime

- Mapping:
- PO -> ALL Items Report (internal key `ALL_ITEMS`)
 - EWASTE-PO -> DTSC
 - SB20-PO -> NETCOST

IMPORTED-DATA:
- Imported tab names are written to A2:A31
- One row per imported invoice tab
