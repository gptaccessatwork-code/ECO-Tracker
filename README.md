# ECO-Tracker

This script appends rows from a local Excel file into an existing Excel table in a target workbook.

Recommended if you do not have Entra app registration access:
- Sync the SharePoint workbook locally with OneDrive
- Point the script at that local synced file
- Let OneDrive upload the saved changes back to SharePoint

## Install

```powershell
pip install -r requirements.txt
```

## Required values

You can pass these as command-line arguments or environment variables.

- `TARGET_WORKBOOK_FILE` or `WORKBOOK_URL` or `WORKBOOK_PATH`
- `TABLE_NAME` like `Table1`
- `EXCEL_FILE`

Optional:

- `TENANT_ID`
- `CLIENT_ID`
- `SHAREPOINT_HOSTNAME` like `contoso.sharepoint.com`
- `SHAREPOINT_SITE_PATH` like `/sites/ECO-Tracker`
- `EXCEL_SHEET`
- `BATCH_SIZE`
- `TOKEN_CACHE_FILE`
- `DRY_RUN=true`

## Run

Using a synced local workbook is easiest if you do not have admin/app-registration access:

```powershell
python .\excel_to_sharepoint.py `
  --target-workbook-file "C:\Users\kmageshkumar\OneDrive - Ichor Systems\AMAT SGP ECO Tracker.xlsx" `
  --table-name "Table1" `
  --excel-file "C:\Users\kmageshkumar\Downloads\AMAT SGP ECO Tracker.xlsx"
```

Using the workbook URL is useful if Graph access is available:

```powershell
python .\excel_to_sharepoint.py `
  --workbook-url "https://contoso.sharepoint.com/:x:/r/personal/user_contoso_com/_layouts/15/Doc.aspx?..." `
  --table-name "Table1" `
  --excel-file ".\data.xlsx"
```

You can also use the SharePoint path if you know it:

```powershell
python .\excel_to_sharepoint.py --workbook-path "Shared Documents/Tracker.xlsx" --table-name "Table1" --excel-file ".\data.xlsx"
```

If your Excel headers do not match the target workbook table columns, add mappings:

```powershell
python .\excel_to_sharepoint.py `
  --workbook-path "Shared Documents/Tracker.xlsx" `
  --table-name "Table1" `
  --excel-file ".\data.xlsx" `
  --map "LocalHeader1=SharePointColumn1" `
  --map "LocalHeader2=SharePointColumn2"
```

## Important note

If you use `TARGET_WORKBOOK_FILE`, the script edits the local synced file directly and does not use Microsoft Graph.

If you use `WORKBOOK_URL` or `WORKBOOK_PATH`, the script uses Microsoft Graph Excel APIs and the first run will prompt you to sign in with device code flow.
