These files are updated Outlook VBA modules based on your current `v3` logic.

What they add:
- Keep the current CRF and Spec Award detection flow
- Queue CRF mail into a Downloads staging workbook, then sync it into the OneDrive workbook when the target is free
- Parse the latest Spec Award email table from the HTML body
- Extract only `Slot Number` and `Part Number`
- Build `System Number` by replacing the `SEMSYS` or `SEMNSO` placeholder with the slot number
- Append new extracted rows into the local staging workbook queue
- Store only `System Number` and the received `Spec Award Date`
- Run `excel_to_sharepoint.py` to sync those rows into the synced SharePoint workbook
- If the synced workbook is busy, keep the queued rows and retry every 15 minutes until the sync succeeds

Files:
- `ThisOutlookSession.bas`
- `Module1.bas`

Config values to check in `Module1.bas`:
- `SPEC_AWARD_STAGING_FILE`
- `SPEC_AWARD_STAGING_SHEET`
- `SYNC_PYTHON_COMMAND`
- `SYNC_SCRIPT_PATH`
- `SYNC_TARGET_WORKBOOK_FILE`
- `SYNC_TARGET_TABLE_NAME`
- `CRF_STAGING_FILE`
- `CRF_STAGING_SHEET`
- `CRF_TARGET_WORKBOOK_FILE`
- `CRF_TARGET_SHEET`
- `CRF_SCRIPT_PATH`

Recommended flow:
1. Close the synced target workbook when possible so the first sync attempt can finish immediately.
2. Import these two `.bas` files into your Outlook VBA project.
3. Restart Outlook or run `StartCRFMonitoring`.
4. Send yourself a Spec Award test email that contains the target table.
5. Check `Documents\OutlookMacroLog.txt` if the extraction or sync fails.

Important note:
The staging workbook now acts as a pending queue. New Spec Award rows are appended there, and only the rows that have been successfully pushed are removed from the staging workbook.
