# Power Automate migration plan

## Recommended destination

The final cloud workflow should be:

```text
Outlook Online
  -> Power Automate
  -> Excel Online (Business): Run script
```

This is the active migration path when Power Automate Premium is unavailable.
The Outlook and Excel Online (Business) connectors are Standard connectors.

The previously considered HTTP bridge requires capabilities that are not
available with the current license:

```text
Outlook VBA
  -> local staging workbook
  -> Python
  -> Power Automate HTTP trigger
  -> Excel Online (Business)
```

Do not use that HTTP path. Outlook and Excel already have Standard Power
Automate cloud connectors, so the final version does not need Outlook desktop,
VBA, a local workbook, Python, an HTTP endpoint, or a permanently running PC.

## What the current solution does

### CRF intake

1. Watch the Inbox for mail from `donotreply@amat.com`.
2. Ignore a message if the Change Coordinator is already a recipient.
3. Extract a CRF number from the subject using:
   `CRF\s*[:#-]?\s*(\d{5,})`.
4. Skip CRFs already processed.
5. Forward the message and mark related CRF messages as read.
6. Append `CRF Number` and `Received Date` to `Table1`.

### Spec Award intake

1. Watch for approved senders and `spec award` in the subject.
2. Require `part number`, `qtr`, and either `need by date` or `nbd` in the
   newest part of the message body.
3. Read `Slot Number` and `Part Number` from the HTML table.
4. Replace the first `SEMSYS` or `SEMNSO` in the part number with the slot
   number to create `System Number`.
5. Append `System Number` and `Spec Award Date` to `Table1`.
6. Forward the message.

### Scheduled tracker maintenance

`agile_eco_dates.py` is a separate concern. It calls an internal Agile WSDL,
updates Submitted/Released dates, calculates elapsed days, highlights overdue
items, and sends reminders. Keep this Python job unchanged until email intake is
stable in Power Automate. The internal HTTP WSDL might require an on-premises
data gateway if it is later moved to a cloud flow.

## Migration order

1. **Create a Standard-only Spec Award cloud flow.**
   Use Outlook `When a new email arrives (V3)` and Excel Online (Business)
   `Run script`.
2. **Move the existing parsing into an Office Script.**
   Pass the email HTML and received timestamp into the script. The script parses
   the table, builds System Numbers, deduplicates, and appends to `Table1`.
3. **Add forwarding and mark-as-read actions.**
   Run these only after the workbook update succeeds.
4. **Migrate CRF separately if required.**
   CRF uses a different workbook and must not share the Spec Award intake flow.
5. **Evaluate the Agile scheduled job.**
   Keep Python if the WSDL is reachable only on the corporate network. Consider
   a gateway or a small internal API only if there is a clear operational
   benefit.

## Prerequisites

- Put each target workbook in OneDrive for Business or a SharePoint document
  library.
- Confirm the target range is an actual Excel table named `Table1`.
- Keep the existing column names:
  - CRF: `CRF Number`, `Received Date`
  - ECO: `System Number`, `Spec Award Date`
- Confirm that the flow owner can edit both workbooks.
- Check whether `When an HTTP request is received` is allowed and licensed in
  the tenant. If it is unavailable, skip the HTTP pilot and start with the
  Outlook cloud trigger.
- Decide who will own the production flow. Prefer a service account or at least
  add a co-owner so the automation does not depend on one employee account.

## Standard-only Spec Award flow

Create an **Automated cloud flow** named `ECO - Spec Award intake`.

Use Office 365 Outlook `When a new email arrives (V3)`:

- Folder: Inbox
- Subject Filter: `spec award`
- Include Attachments: No

Add conditions for the approved sender list and required body keywords. Then run
an Office Script against `AMAT SGP ECO Tracker.xlsx`, passing the HTML body and
received timestamp as parameters.

Install `power_automate/spec_award_office_script.ts` from this folder through
Excel for the web: **Automate > New Script > Create in Code Editor**. Name the
saved script `ECO - Process Spec Award`.

The Office Script is responsible for:

- locating an HTML table with `Slot Number` and `Part Number` headers
- replacing the first `SEMSYS` or `SEMNSO` in a part number with the slot number
- skipping duplicate `System Number` values
- appending `System Number` and `Spec Award Date` to `Table1`

## Unused Premium HTTP design

The following section is retained only as a design reference. It is not the
chosen implementation because the account does not have Power Automate Premium.

Create an **Instant cloud flow** named `ECO - Spec Award intake - HTTP`.

### 1. Trigger

Use `When an HTTP request is received`.

Paste the contents of
`power_automate/spec-award-request.schema.json` into
**Request Body JSON Schema**.

For authentication:

- Prefer `Specific users in my tenant` and an approved service principal.
- `Any user in my tenant` also requires the Python sender to obtain an Entra
  access token.
- Use `Anyone` only for a short controlled pilot if company policy permits it.
  In that mode, the generated URL is a secret. Never commit it to this folder or
  put it in a log.

Save the flow once to generate its endpoint URL.

### 2. Validate the event

Add a condition:

```text
eventType is equal to spec-award.received
```

In the false branch, return HTTP `400` with:

```json
{"ok": false, "error": "Unsupported eventType"}
```

### 3. Serialize workbook writes

Open the trigger settings, enable **Concurrency Control**, and set the degree of
parallelism to `1`. This reduces duplicate and workbook-lock problems.

### 4. Process the items

Add `Apply to each` over `items`.

For each item:

1. Use Excel Online (Business) `List rows present in a table`.
2. Select the ECO workbook and `Table1`.
3. Filter for the current `System Number`.
4. Add a condition that the returned `value` array has length `0`.
5. In the yes branch, use `Add a row into a table`:
   - `System Number` = `systemNumber`
   - `Spec Award Date` = `specAwardDate`
6. In the no branch, do nothing; a duplicate is a successful no-op.

Keep `Apply to each` sequential as well.

### 5. Return success only after Excel finishes

Add a `Response` action at the very end:

- Status: `200`
- Body:

```json
{"ok": true, "eventId": "<eventId from trigger>"}
```

Python must remove rows from the local staging queue only after this `200`
response. A timeout, `429`, or `5xx` must leave the rows queued for retry.

## Request contract

CRF example:

```json
{
  "schemaVersion": "1.0",
  "eventType": "crf.received",
  "eventId": "outlook-message-id-or-stable-generated-id",
  "sentAt": "2026-07-28T08:30:00+08:00",
  "items": [
    {
      "crfNumber": "123456",
      "receivedDate": "2026-07-28"
    }
  ]
}
```

Spec Award example:

```json
{
  "schemaVersion": "1.0",
  "eventType": "spec-award.received",
  "eventId": "outlook-message-id-or-stable-generated-id",
  "sentAt": "2026-07-28T08:30:00+08:00",
  "items": [
    {
      "systemNumber": "1234-56789",
      "specAwardDate": "2026-07-28"
    }
  ]
}
```

Use ISO 8601 for timestamps and `yyyy-MM-dd` for date-only fields. Do not send
locale-formatted dates such as `7/8/2026`, because they are ambiguous.

## Idempotency and retries

HTTP delivery is at-least-once, not exactly-once. A request can finish in Excel
while the caller times out, causing the same request to be sent again.

Use two levels of deduplication:

- CRF rows: `CRF Number`
- Spec Award rows: `System Number`

For stronger auditing, add these columns to both tables later:

- `Event ID`
- `Source Message ID`
- `Processed At`

Then use `Event ID` as the request-level idempotency key and retain the existing
business-key duplicate check.

## Cutover test

Run the HTTP pilot against a copied workbook first.

Test all of these:

1. One valid CRF creates one row.
2. Sending the identical payload twice still leaves one row.
3. Two different CRFs both appear.
4. An invalid event type returns `400`.
5. A deliberately unavailable/locked workbook does not delete the local queue.
6. Dates remain correct in Singapore time.
7. Flow run history contains no credentials, endpoint URLs, or sensitive email
   content.

After the copied-workbook tests pass, run both old and new paths in comparison
mode for several days, but allow only one path to write to the production table.

## Operational cautions

- Avoid editing the workbook in Excel Desktop while a flow is writing to it.
- Excel Online is suitable for this small tracker but is not a transactional
  database. If volume or concurrent writers grow, use a SharePoint List or
  Dataverse as the system of record and keep Excel as a report.
- Add failure notification to the flow using a separate failure scope configured
  with `run after`.
- Export the flow in a solution and use environment variables/connection
  references before calling it production-ready.
