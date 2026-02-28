# DTSC Report Wizard

This Google Sheets Apps Script helps generate DTSC report files and import invoice tab data automatically.

## What this tool does

1. Finds the customer folder in Purchases.
2. Creates/uses a Generated Reports folder.
3. Copies report templates:
   - DTSC-ALL
   - DTSC
   - NETCOST (only when SB-20 is checked)
4. Finds invoice files by invoice folder name.
5. Ensures invoice has a PO tab.
6. Imports/refreshes:
   - EWASTE-PO
   - SB20-PO (only when SB-20 is checked)
7. Freezes those tabs to values.
8. Sends tabs into generated report files using this mapping:
   - PO -> DTSC-ALL
   - EWASTE-PO -> DTSC
   - SB20-PO -> NETCOST
9. Updates IMPORTED-DATA in each report:
   - Writes imported tab names in A2:A31
   - Keeps multiple invoice rows (appends up to 30)

## How to run

1. Open the Google Sheet with this script attached.
2. Refresh the sheet.
3. Go to Reports -> Open Progress Window.
4. Enter:
   - Customer Number
   - Invoice Number(s) (comma or new line)
   - SB-20 checkbox
   - Mode (TEST or LIVE)
5. Click Start Run.

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

## Copy/paste summary for Google Sheet tab

Use this short version in a Sheet note/tab:

- Open Reports -> Open Progress Window
- Enter Customer Number and Invoice Number(s)
- Check SB-20 if needed
- Select TEST (safe) or LIVE (real updates)
- Click Start Run and watch progress/logs

Mapping:
- PO -> DTSC-ALL
- EWASTE-PO -> DTSC
- SB20-PO -> NETCOST

IMPORTED-DATA:
- Imported tab names are written to A2:A31
- One row per imported invoice tab
