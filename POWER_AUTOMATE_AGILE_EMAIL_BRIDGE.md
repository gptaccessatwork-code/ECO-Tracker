# ECO Tracker Agile email bridge migration (Phase 1)

This design removes OneDrive synchronization from the scheduled tracker update.
Power Automate reads and writes the SharePoint ECO workbook with Excel Online.
The local Python task only receives JSON through Outlook, calls the internal
Agile SOAP service, and returns JSON through Outlook.

## Components supplied in this folder

- `agile_eco_dates.py`: run with `--email-bridge`
- `power_automate/agile_export_pending_office_script.ts`
- `power_automate/agile_import_results_office_script.ts`
- `power_automate/agile-request.schema.json`
- `power_automate/agile-result.schema.json`

Phase 1 covers only `AMAT SGP ECO Tracker.xlsx`. CRF support in the code is
reserved for a future phase and must not be configured in either flow yet. The
existing local-workbook mode remains available for rollback. Do not run the old
ECO workbook-writing task and the new ECO import flow against production at the
same time.

## Before building the flows

1. Confirm the SharePoint ECO workbook contains an Excel table named `Table1`.
2. Confirm the ECO table contains these headers:
   `System Number`, `Spec Award Date`, `ECO Number`, `Submitted Date`,
   `Delta (Spec to Submitted)`, `Delta (excl weekend)`, `Released Date`,
   `Delta (Spec to Released)`, and `Delta (excl weekend)2`.
3. Make a copy of the ECO workbook in SharePoint for the first end-to-end test.
4. Decide which work mailbox owns the flows. Use the same SMTP address for the
   scheduled-flow sender, job recipient, result sender, and result recipient.

The current ECO worksheet is protected without a password. The import script
temporarily pauses that protection while updating `Table1` and resumes it in a
`finally` block. If password protection is added later, the import will stop
with an explicit error instead of silently failing or weakening protection.

## Install the Office Scripts

In Excel for the web, open a workbook and select **Automate > New Script >
Create in Code Editor**.

1. Paste `agile_export_pending_office_script.ts` and save it as
   `Agile - Export Pending Rows`.
2. Paste `agile_import_results_office_script.ts` and save it as
   `Agile - Import Results`.

If scripts are saved in OneDrive, use the Excel Online (Business) **Run script**
action. If the `.osts` files are saved in a team SharePoint library, use **Run
script from SharePoint library**.

## Flow 1: create the ECO Agile lookup job

Create a **Scheduled cloud flow** named `Agile Tracker - Create Lookup Jobs`.

### Trigger

- Frequency: `Day`
- Interval: `1`
- Time zone: `Singapore Standard Time`
- At these hours: `5`
- At these minutes: `45`

Open trigger settings, enable concurrency control, and set degree of
parallelism to `1`.

1. Add **Compose** named `ECO Job ID` with this expression:

   ```text
   concat(formatDateTime(convertTimeZone(utcNow(),'UTC','Singapore Standard Time'),'yyyyMMdd-HHmmss'),'-ECO')
   ```

2. Add Excel Online (Business) **Run script** named `Export ECO rows`:
   - Location: the ECO workbook's SharePoint site
   - Document Library: the workbook library
   - File: select the ECO workbook with the file picker
   - Script: `Agile - Export Pending Rows`
   - `trackerType`: `ECO`
   - `jobId`: output of `ECO Job ID`
   - `tableName`: `Table1`

3. Add Office 365 Outlook **Send an email (V2)**:
   - To: the bridge mailbox
   - Subject: `AGILE-JOB | ` followed by the ECO Job ID
   - Body: `Automated Agile ECO lookup request.`
   - Attachment Name:

     ```text
     concat('agile-job-',outputs('ECO_Job_ID'),'.json')
     ```

   - Attachment Content:

     ```text
     base64(outputs('Export_ECO_rows')?['body/result'])
     ```

Action names in expressions must match the internal names Power Automate shows.
If the designer renamed an action, select the action output from Dynamic content
or update the expression name.

Do not add a CRF branch in Phase 1. Keep this flow pointed at the copied ECO
workbook until cutover.

## Configure and test the local worker

The `agile_eco_dates.exe` in this folder has been rebuilt with email-bridge
support, so the pilot and scheduled task do not depend on a Python installation.

Use this non-destructive self-test first:

```powershell
.\agile_eco_dates.exe `
  --bridge-self-test `
  --bridge-mailbox "your.email@ichorsystems.com" `
  --bridge-job-sender "your.email@ichorsystems.com" `
  --bridge-folder "Inbox"
```

The test checks the mailbox configuration, Agile credential file, Outlook COM,
and source folder. It does not read, move, or send messages.

Process waiting jobs manually during the pilot:

```powershell
.\agile_eco_dates.exe `
  --email-bridge `
  --bridge-mailbox "your.email@ichorsystems.com" `
  --bridge-job-sender "your.email@ichorsystems.com" `
  --bridge-folder "Inbox"
```

The worker launches Classic Outlook through COM if it is not already running,
starts Outlook Send/Receive, and waits up to five minutes for the cloud mailbox
to synchronize. It scans every 10 seconds and exits as soon as it handles a
job. To change those defaults, add `--bridge-wait-seconds 300` and
`--bridge-poll-seconds 10`. Use `--bridge-wait-seconds 0` only when an immediate
single scan is desired.

On success, Python sends `AGILE-RESULT | <jobId>`, records the job in
`.agile_email_bridge_state.json`, and moves the job message into the
`Processed` subfolder under the configured Outlook folder. Failed jobs stay in
the source folder for retry.

## Flow 2: import Agile results

Create an **Automated cloud flow** named `Agile Tracker - Import Results`.

### Trigger

Use Office 365 Outlook **When a new email arrives (V3)**:

- Folder: `Inbox`
- From: the bridge mailbox SMTP address
- Subject Filter: `AGILE-RESULT |`
- Only with Attachments: `Yes`
- Include Attachments: `No`

Use the trigger filters rather than a later Condition so unrelated email does
not consume flow runs.

Open trigger settings, enable concurrency control, and set degree of
parallelism to `1` so retried result emails cannot write simultaneously.

### Read and validate the result

1. Add **Apply to each** over the trigger's `Attachments` array.
2. Add a Condition inside it:

   ```text
   startsWith(toLower(items('Apply_to_each')?['name']),'agile-result-')
   ```

3. In the Yes branch, add Outlook **Get Attachment (V2)**:
   - Message Id: trigger `Message Id`
   - Attachment Id: current attachment `Id`
4. Add **Compose** named `Decoded result JSON`:

   ```text
   base64ToString(outputs('Get_Attachment_(V2)')?['body/contentBytes'])
   ```

5. Add **Parse JSON**:
   - Content: output of `Decoded result JSON`
   - Schema: paste `power_automate/agile-result.schema.json`
6. Add a Condition: parsed `trackerType` is equal to `ECO`. Terminate the flow
   as failed when this condition is false; CRF results are not accepted in
   Phase 1.

In the Yes branch, run `Agile - Import Results` against the ECO workbook:

- `resultJson`: output of `Decoded result JSON`
- `tableName`: `Table1`

The import script:

- validates the payload
- checks that the source row still contains the expected identifiers
- falls back to identifier matching if the table was sorted
- never clears existing dates when an Agile lookup failed
- calculates the four delta columns
- handles blank and not-applicable tracking numbers
- returns reminder rows and lookup errors
- records the `jobId` in a hidden `_AgileAutomationLog` sheet
- treats a repeated result email as a successful no-op

### Reminder action

After the **Run script** import action, add a Condition:

```text
greater(length(outputs('Import_ECO_results')?['body/result/reminders']),0)
```

In the Yes branch:

1. Add **Create HTML table** using the returned `reminders` array.
2. Add Outlook **Send an email (V2)** with the HTML table in the body.
3. Use the production To/CC/BCC values from the old scheduled task.

Also add an operator alert when `failedLookups` or `unmatchedRows` is greater
than zero. Failed Agile lookups are reported but do not erase workbook dates.

## Task Scheduler cutover

Do this only after the copied-workbook test completes successfully.

Edit the existing 6:00 AM task:

- Program/script:
  `C:\Users\kmageshkumar\OneDrive - Ichor Systems\Scripts\ECO & CRF Tracker\agile_eco_dates.exe`
- Add arguments:

  ```text
  --email-bridge --bridge-mailbox "your.email@ichorsystems.com" --bridge-job-sender "your.email@ichorsystems.com" --bridge-folder "Inbox"
  ```

- Start in:
  `C:\Users\kmageshkumar\OneDrive - Ichor Systems\Scripts\ECO & CRF Tracker`
- Select **Run only when user is logged on** because Outlook COM requires the
  interactive Outlook profile.
- Select **Run task as soon as possible after a scheduled start is missed**.
- A separate task to open Outlook is not required. The worker launches Classic
  Outlook, requests Send/Receive, and waits for mailbox synchronization.
- Optionally repeat every `15 minutes` for `1 hour` as a second layer of
  protection against an unusually long Outlook or network delay.
- If the task is already running: **Do not start a new instance**.

The task exits after processing the available messages or after its five-minute
wait expires; it does not stay running continuously. Classic Outlook may remain
open after the task finishes. New Outlook cannot provide this COM interface.

## End-to-end acceptance test

1. Point both flows to the copied SharePoint ECO workbook.
2. Manually run Flow 1.
3. Confirm one `AGILE-JOB` message arrives with a valid JSON attachment.
4. Run `agile_eco_dates.exe` manually in `--email-bridge` mode.
5. Confirm one `AGILE-RESULT` message arrives.
6. Confirm Flow 2 succeeds and updates only the copied ECO workbook.
7. Resend one result email and verify `duplicateJob` is `true` and no duplicate
   reminder is sent.
8. Temporarily alter a copied row after export and verify it is either found by
   its identifier or reported as unmatched, never written to the wrong row.
9. Confirm lookup failures appear in the operator alert and do not blank dates.
10. Switch both flows to the production ECO workbook.
11. Change the Task Scheduler action to email-bridge mode.
12. Disable the old local-workbook-writing action. Keep its configuration for
    rollback until several successful production runs have completed.

## Rollback

Disable both new flows, restore the previous Task Scheduler arguments that pass
`--workbook-file`, and run the old ECO mode once. Do not enable both ECO writers
concurrently.

## CRF during Phase 1

The new Power Automate flows do not read or write the CRF workbook. If the CRF
workbook still requires its existing local update, keep that work in a separate
Task Scheduler action and explicitly pass an empty ECO workbook value so it
cannot touch the ECO tracker:

```text
--workbook-file "" --crf-workbook-file "C:\Users\kmageshkumar\OneDrive - Ichor Systems\AMAT SGP CRF Tracker.xlsx"
```

Do not send `trackerType=CRF` email jobs until the separate CRF migration is
tested and approved.
