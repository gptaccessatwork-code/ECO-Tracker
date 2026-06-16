# ECO Tracker Scripts

This folder contains two main scripts:

- `excel_to_sharepoint.py`
  - Appends rows from a local Excel file into an existing Excel table in a target workbook
  - Can write either to a locally synced OneDrive/SharePoint workbook or through Microsoft Graph
- `agile_eco_dates.py`
  - Looks up ECO workflow dates from Agile WSDL
  - Updates the workbook with Submitted Date, Released Date, delta columns, conditional highlighting, and reminder emails

Recommended if you do not have Entra app registration access:

- Sync the SharePoint workbook locally with OneDrive
- Point the scripts at that local synced file
- Let OneDrive upload the saved changes back to SharePoint

## Install

```powershell
pip install -r requirements.txt
```

## excel_to_sharepoint.py

### Required values

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

### Run

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

### Important note

If you use `TARGET_WORKBOOK_FILE`, the script edits the local synced file directly and does not use Microsoft Graph.

If you use `WORKBOOK_URL` or `WORKBOOK_PATH`, the script uses Microsoft Graph Excel APIs and the first run will prompt you to sign in with device code flow.

### Conflict retry queue

For the Spec Award intake flow, `excel_to_sharepoint.py` can now be used in queue mode:

- the source staging workbook acts as the pending queue
- if the target workbook is busy, a background retry worker is launched
- the worker retries every 15 minutes until the write succeeds
- only the processed rows are removed from the staging workbook after a successful sync
- rows that arrive while a retry is waiting remain in the staging workbook and are included in a later successful sync

## agile_eco_dates.py

### Purpose

This script updates the tracker workbook in place by:

- reading ECO numbers from column `C`
- writing Submitted Date to column `D`
- writing Released Date to column `G`
- calculating delta columns
- applying overdue highlighting
- sending reminder emails for overdue rows

Default workbook columns:

- `A`: System Number
- `B`: Spec Award Date
- `C`: ECO Number
- `D`: Submitted Date
- `E`: Delta (Spec to Submitted)
- `F`: Delta (excl weekend)
- `G`: Released Date
- `H`: Delta (Spec to Released)
- `I`: Delta (excl weekend)2

### Required setup

Environment variables:

- `AGILE_USER`
- `AGILE_PASS`

Optional environment variables:

- `TARGET_WORKBOOK_FILE`
- `TARGET_WORKSHEET`
- `ECO_REMINDER_TO`
- `ECO_REMINDER_CC`
- `ECO_REMINDER_BCC`

### Run

Update the workbook and send reminders if overdue rows are found:

```powershell
python .\agile_eco_dates.py `
  --workbook-file "C:\Users\kmageshkumar\OneDrive - Ichor Systems\AMAT SGP ECO Tracker.xlsx" `
  --reminder-to "your.email@ichorsystems.com"
```

Check a single ECO only:

```powershell
python .\agile_eco_dates.py DSM25859
```

### Reminder rules

The script flags and follows up on these conditions:

- no ECO number and Spec Award Date is 2 or more days old
- ECO exists but is still not submitted 2 or more days after Spec Award Date
- Submitted Date exists but Released Date is still blank 3 or more days later

### Testing modes

Run a setup-only self-test without modifying the workbook or sending email:

```powershell
python .\agile_eco_dates.py `
  --workbook-file "C:\Users\kmageshkumar\OneDrive - Ichor Systems\AMAT SGP ECO Tracker.xlsx" `
  --self-test
```

This checks items such as:

- `AGILE_USER`
- `AGILE_PASS`
- Outlook COM availability
- workbook file existence
- worksheet availability
- reminder recipient configuration

Preview the reminder email as an Outlook draft instead of sending it:

```powershell
python .\agile_eco_dates.py `
  --workbook-file "C:\Users\kmageshkumar\OneDrive - Ichor Systems\AMAT SGP ECO Tracker.xlsx" `
  --preview-reminder `
  --reminder-to "your.email@ichorsystems.com"
```

Notes:

- `--preview-reminder` still updates the workbook
- the email opens as a draft window using Outlook instead of sending immediately
- preview mode does not mark the reminder as sent in `.eco_reminder_state.json`
