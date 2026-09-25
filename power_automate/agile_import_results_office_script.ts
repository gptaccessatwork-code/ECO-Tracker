interface AgileLookupResultItem {
  recordType: string;
  recordId: string;
  sourceRow?: number;
  systemNumber?: string;
  trackingNumber: string;
  baseDate?: string | null;
  lookupSucceeded: boolean;
  lookupStatus: string;
  lookupError?: string;
  submittedDate?: string | null;
  releasedDate?: string | null;
}

interface ReminderRow {
  recordType: string;
  recordId: string;
  systemNumber: string;
  trackingNumber: string;
  baseDate: string;
  submittedDate: string;
  releasedDate: string;
  reason: string;
}

interface ImportSummary {
  duplicateJob: boolean;
  jobId: string;
  trackerType: string;
  updatedRows: number;
  failedLookups: number;
  unmatchedRows: number;
  reminders: ReminderRow[];
  errors: string[];
}

interface TrackerColumnIndexes {
  systemNumber: number;
  tracking: number;
  baseDate: number;
  submitted: number;
  submittedDelta: number;
  submittedDeltaNoWeekend: number;
  released: number;
  releasedDelta: number;
  releasedDeltaNoWeekend: number;
}

/**
 * Apply a JSON result attachment produced by agile_eco_dates.py.
 * Power Automate should pass the decoded attachment text as resultJson.
 */
function main(
  workbook: ExcelScript.Workbook,
  resultJson: string,
  tableName: string = "Table1"
): ImportSummary {
  const payload = parsePayload(resultJson);
  const jobId = String(payload.jobId);
  const trackerType = String(payload.trackerType).toUpperCase();
  const items = payload.items as AgileLookupResultItem[];
  const emptySummary: ImportSummary = {
    duplicateJob: false,
    jobId,
    trackerType,
    updatedRows: 0,
    failedLookups: 0,
    unmatchedRows: 0,
    reminders: [],
    errors: []
  };

  if (hasProcessedJob(workbook, jobId)) {
    emptySummary.duplicateJob = true;
    return emptySummary;
  }

  const table = findTable(workbook, tableName);
  const protection = table.getWorksheet().getProtection();
  let protectionPaused = false;
  if (protection.getProtected() && !protection.getIsPaused()) {
    if (!protection.getCanPauseProtection()) {
      throw new Error("The ECO worksheet is protected and protection cannot be paused.");
    }
    if (protection.getIsPasswordProtected()) {
      throw new Error(
        "The ECO worksheet uses password protection. Configure editable table cells or remove the password before using this flow."
      );
    }
    protection.pauseProtection();
    protectionPaused = true;
  }

  try {
  const headers = table.getHeaderRowRange().getTexts()[0];
  const indexes = getColumnIndexes(headers, trackerType);
  if (table.getRowCount() === 0) {
    recordProcessedJob(workbook, jobId, trackerType, 0);
    return emptySummary;
  }

  const body = table.getRangeBetweenHeaderAndTotal();
  const values = body.getValues();
  const texts = body.getTexts();
  const today = singaporeToday();

  for (const item of items) {
    if (String(item.recordType).toUpperCase() !== trackerType) {
      emptySummary.errors.push(`${item.recordId}: result has the wrong recordType.`);
      continue;
    }

    const rowIndex = findTargetRow(item, values, texts, indexes, trackerType);
    if (rowIndex < 0) {
      emptySummary.unmatchedRows += 1;
      emptySummary.errors.push(`${item.recordId}: matching workbook row was not found.`);
      continue;
    }

    const trackingNumber = texts[rowIndex][indexes.tracking].trim();
    const baseDate = parseDateCell(values[rowIndex][indexes.baseDate], texts[rowIndex][indexes.baseDate]);
    const notApplicable = isNotApplicable(trackingNumber);

    let submittedDate = parseDateCell(
      values[rowIndex][indexes.submitted],
      texts[rowIndex][indexes.submitted]
    );
    let releasedDate = parseDateCell(
      values[rowIndex][indexes.released],
      texts[rowIndex][indexes.released]
    );

    if (notApplicable) {
      values[rowIndex][indexes.submittedDelta] = "";
      values[rowIndex][indexes.submittedDeltaNoWeekend] = "";
      values[rowIndex][indexes.releasedDelta] = "";
      values[rowIndex][indexes.releasedDeltaNoWeekend] = "";
      formatDeltaCell(body, rowIndex, indexes.submittedDelta, false);
      formatDeltaCell(body, rowIndex, indexes.submittedDeltaNoWeekend, false);
      formatDeltaCell(body, rowIndex, indexes.releasedDelta, false);
      formatDeltaCell(body, rowIndex, indexes.releasedDeltaNoWeekend, false);
      emptySummary.updatedRows += 1;
      continue;
    }

    if (item.lookupSucceeded && item.lookupStatus === "completed") {
      submittedDate = parseIsoDate(item.submittedDate || null);
      releasedDate = parseIsoDate(item.releasedDate || null);
      values[rowIndex][indexes.submitted] = submittedDate ? toExcelSerial(submittedDate) : "";
      values[rowIndex][indexes.released] = releasedDate ? toExcelSerial(releasedDate) : "";
    } else if (!item.lookupSucceeded) {
      emptySummary.failedLookups += 1;
      emptySummary.errors.push(
        `${item.recordId}: ${String(item.lookupError || "Agile lookup failed.").slice(0, 500)}`
      );
    }

    const submittedEnd = submittedDate || today;
    const releasedEnd = releasedDate || today;
    values[rowIndex][indexes.submittedDelta] = dateDelta(baseDate, submittedEnd);
    values[rowIndex][indexes.submittedDeltaNoWeekend] = businessDateDelta(baseDate, submittedEnd);
    values[rowIndex][indexes.releasedDelta] = dateDelta(baseDate, releasedEnd);
    values[rowIndex][indexes.releasedDeltaNoWeekend] = businessDateDelta(baseDate, releasedEnd);

    const submittedAlert =
      !submittedDate &&
      typeof values[rowIndex][indexes.submittedDelta] === "number" &&
      Number(values[rowIndex][indexes.submittedDelta]) >= 2;
    const releasedAlert =
      Boolean(submittedDate) &&
      !releasedDate &&
      Number(dateDelta(submittedDate, today)) >= 3;
    formatDeltaCell(body, rowIndex, indexes.submittedDelta, submittedAlert);
    formatDeltaCell(body, rowIndex, indexes.submittedDeltaNoWeekend, false);
    formatDeltaCell(body, rowIndex, indexes.releasedDelta, releasedAlert);
    formatDeltaCell(body, rowIndex, indexes.releasedDeltaNoWeekend, false);

    const reason = followUpReason(
      trackerType,
      baseDate,
      trackingNumber,
      submittedDate,
      releasedDate,
      today
    );
    if (reason && baseDate) {
      emptySummary.reminders.push({
        recordType: trackerType,
        recordId: item.recordId,
        systemNumber:
          trackerType === "ECO" ? texts[rowIndex][indexes.systemNumber].trim() : "",
        trackingNumber,
        baseDate: formatDate(baseDate),
        submittedDate: submittedDate ? formatDate(submittedDate) : "",
        releasedDate: releasedDate ? formatDate(releasedDate) : "",
        reason
      });
    }
    emptySummary.updatedRows += 1;
  }

  writeColumn(table, headers, indexes.submitted, values);
  writeColumn(table, headers, indexes.submittedDelta, values);
  writeColumn(table, headers, indexes.submittedDeltaNoWeekend, values);
  writeColumn(table, headers, indexes.released, values);
  writeColumn(table, headers, indexes.releasedDelta, values);
  writeColumn(table, headers, indexes.releasedDeltaNoWeekend, values);
  table.getColumnByName(headers[indexes.submitted]).getRangeBetweenHeaderAndTotal().setNumberFormat("dd mmm yyyy");
  table.getColumnByName(headers[indexes.released]).getRangeBetweenHeaderAndTotal().setNumberFormat("dd mmm yyyy");
  recordProcessedJob(workbook, jobId, trackerType, emptySummary.updatedRows);
  return emptySummary;
  } finally {
    if (protectionPaused) {
      protection.resumeProtection();
    }
  }
}

function parsePayload(resultJson: string): { [key: string]: unknown } {
  let payload: { [key: string]: unknown };
  try {
    payload = JSON.parse(resultJson) as { [key: string]: unknown };
  } catch (_error) {
    throw new Error("resultJson is not valid JSON.");
  }
  if (payload.schemaVersion !== "1.0" || payload.eventType !== "agile.lookup.result") {
    throw new Error("Unexpected Agile result schemaVersion or eventType.");
  }
  if (!payload.jobId || !Array.isArray(payload.items)) {
    throw new Error("Agile result is missing jobId or items.");
  }
  const trackerType = String(payload.trackerType || "").toUpperCase();
  if (trackerType !== "ECO" && trackerType !== "CRF") {
    throw new Error("Agile result trackerType must be ECO or CRF.");
  }
  return payload;
}

function getColumnIndexes(headers: string[], trackerType: string): TrackerColumnIndexes {
  return {
    systemNumber: trackerType === "ECO" ? requireHeader(headers, ["System Number"]) : -1,
    tracking: requireHeader(headers, [trackerType === "ECO" ? "ECO Number" : "CRF Number"]),
    baseDate: requireHeader(
      headers,
      trackerType === "ECO"
        ? ["Spec Award Date"]
        : ["CRF Date", "Received Date", "CRF Received Date"]
    ),
    submitted: requireHeader(headers, ["Submitted Date"]),
    submittedDelta: requireHeader(
      headers,
      trackerType === "ECO"
        ? ["Delta (Spec to Submitted)", "Submitted Delta"]
        : ["Delta (CRF to Submitted)", "Delta (Received to Submitted)", "Submitted Delta"]
    ),
    submittedDeltaNoWeekend: requireHeader(
      headers,
      ["Delta (excl weekend)", "Submitted Delta (excl weekend)"]
    ),
    released: requireHeader(headers, ["Released Date"]),
    releasedDelta: requireHeader(
      headers,
      trackerType === "ECO"
        ? ["Delta (Spec to Released)", "Released Delta"]
        : ["Delta (CRF to Released)", "Delta (Received to Released)", "Released Delta"]
    ),
    releasedDeltaNoWeekend: requireHeader(
      headers,
      ["Delta (excl weekend)2", "Released Delta (excl weekend)"]
    )
  };
}

function findTargetRow(
  item: AgileLookupResultItem,
  _values: (string | number | boolean)[][],
  texts: string[][],
  indexes: TrackerColumnIndexes,
  trackerType: string
): number {
  const expectedTracking = normalizeValue(item.trackingNumber);
  const expectedSystem = normalizeValue(item.systemNumber || "");
  const sourceRow = Number(item.sourceRow);
  if (
    Number.isInteger(sourceRow) &&
    sourceRow >= 0 &&
    sourceRow < texts.length &&
    normalizeValue(texts[sourceRow][indexes.tracking]) === expectedTracking &&
    (trackerType !== "ECO" || normalizeValue(texts[sourceRow][indexes.systemNumber]) === expectedSystem)
  ) {
    return sourceRow;
  }

  for (let index = 0; index < texts.length; index += 1) {
    if (trackerType === "ECO" && expectedSystem) {
      if (normalizeValue(texts[index][indexes.systemNumber]) === expectedSystem) {
        return index;
      }
    } else if (normalizeValue(texts[index][indexes.tracking]) === expectedTracking) {
      return index;
    }
  }
  return -1;
}

function writeColumn(
  table: ExcelScript.Table,
  headers: string[],
  columnIndex: number,
  values: (string | number | boolean)[][]
): void {
  table
    .getColumnByName(headers[columnIndex])
    .getRangeBetweenHeaderAndTotal()
    .setValues(values.map((row) => [row[columnIndex]]));
}

function formatDeltaCell(
  body: ExcelScript.Range,
  rowIndex: number,
  columnIndex: number,
  alert: boolean
): void {
  const format = body.getCell(rowIndex, columnIndex).getFormat();
  format.getFill().clear();
  format.getFont().setColor("#000000");
  format.getFont().setBold(alert);
  format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  if (alert) {
    format.getFill().setColor("#FF5B5B");
  }
}

function findTable(workbook: ExcelScript.Workbook, tableName: string): ExcelScript.Table {
  const expected = tableName.trim().toLowerCase();
  const tables = workbook.getTables();
  for (let index = 0; index < tables.length; index += 1) {
    const candidate = tables[index];
    if (candidate.getName().trim().toLowerCase() === expected) {
      return candidate;
    }
  }
  throw new Error(`Could not find Excel table '${tableName}'.`);
}

function requireHeader(headers: string[], acceptedNames: string[]): number {
  const accepted = acceptedNames.map(normalizeHeader);
  const index = headers.findIndex((header) => accepted.includes(normalizeHeader(header)));
  if (index < 0) {
    throw new Error(`Missing required table column: ${acceptedNames.join(" or ")}.`);
  }
  return index;
}

function normalizeHeader(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function normalizeValue(value: string): string {
  return String(value || "").trim().toUpperCase();
}

function isNotApplicable(value: string): boolean {
  const normalized = value.toLowerCase().replace(/[^a-z0-9]/g, "");
  return normalized === "na" || normalized === "none" || normalized === "notapplicable";
}

function parseDateCell(value: string | number | boolean, text: string): Date | null {
  if (typeof value === "number" && isFinite(value) && value > 0) {
    return new Date(Math.round((value - 25569) * 86400000));
  }
  const candidate = String(text || value || "").trim();
  if (!candidate) {
    return null;
  }
  const timestamp = Date.parse(candidate);
  return isNaN(timestamp) ? null : new Date(timestamp);
}

function parseIsoDate(value: string | null): Date | null {
  if (!value) {
    return null;
  }
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) {
    throw new Error(`Invalid ISO date '${value}'.`);
  }
  return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
}

function toExcelSerial(value: Date): number {
  return Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()) / 86400000 + 25569;
}

function singaporeToday(): Date {
  const singapore = new Date(Date.now() + 8 * 60 * 60 * 1000);
  return new Date(Date.UTC(singapore.getUTCFullYear(), singapore.getUTCMonth(), singapore.getUTCDate()));
}

function dateDelta(start: Date | null, end: Date | null): number | "" {
  if (!start || !end) {
    return "";
  }
  return Math.abs(Math.round((dayValue(end) - dayValue(start)) / 86400000));
}

function businessDateDelta(start: Date | null, end: Date | null): number | "" {
  if (!start || !end) {
    return "";
  }
  let first = new Date(dayValue(start));
  let last = new Date(dayValue(end));
  if (first.getTime() > last.getTime()) {
    const swap = first;
    first = last;
    last = swap;
  }
  const raw = Math.round((last.getTime() - first.getTime()) / 86400000);
  let weekends = 0;
  for (
    let current = new Date(first.getTime() + 86400000);
    current.getTime() <= last.getTime();
    current = new Date(current.getTime() + 86400000)
  ) {
    if (current.getUTCDay() === 0 || current.getUTCDay() === 6) {
      weekends += 1;
    }
  }
  return raw - weekends;
}

function dayValue(value: Date): number {
  return Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate());
}

function followUpReason(
  trackerType: string,
  baseDate: Date | null,
  trackingNumber: string,
  submittedDate: Date | null,
  releasedDate: Date | null,
  today: Date
): string {
  if (!baseDate) {
    return "";
  }
  const daysSinceBase = Number(dateDelta(baseDate, today));
  if (!trackingNumber.trim() && daysSinceBase >= 2) {
    return trackerType === "ECO"
      ? "No ECO number after 2 days from Spec Award date"
      : "No CRF number after 2 days from Received date";
  }
  if (!submittedDate && daysSinceBase >= 2) {
    return trackerType === "ECO"
      ? "ECO not submitted within 2 days from Spec Award date"
      : "CRF not submitted within 2 days from Received date";
  }
  if (submittedDate && !releasedDate && Number(dateDelta(submittedDate, today)) >= 3) {
    return trackerType === "ECO"
      ? "ECO not released after 3 days from Submitted date"
      : "CRF not released after 3 days from Submitted date";
  }
  return "";
}

function formatDate(value: Date): string {
  return `${value.getUTCFullYear()}-${String(value.getUTCMonth() + 1).padStart(2, "0")}-${String(
    value.getUTCDate()
  ).padStart(2, "0")}`;
}

function hasProcessedJob(workbook: ExcelScript.Workbook, jobId: string): boolean {
  const sheet = workbook.getWorksheet("_AgileAutomationLog");
  if (!sheet) {
    return false;
  }
  const used = sheet.getUsedRange();
  if (!used) {
    return false;
  }
  return used.getTexts().some((row) => row[0] === jobId);
}

function recordProcessedJob(
  workbook: ExcelScript.Workbook,
  jobId: string,
  trackerType: string,
  updatedRows: number
): void {
  let sheet = workbook.getWorksheet("_AgileAutomationLog");
  if (!sheet) {
    sheet = workbook.addWorksheet("_AgileAutomationLog");
    sheet.getRange("A1:D1").setValues([["Job ID", "Tracker", "Processed At", "Updated Rows"]]);
  }
  const used = sheet.getUsedRange();
  const nextRow = used ? used.getRowCount() : 1;
  sheet.getRangeByIndexes(nextRow, 0, 1, 4).setValues([
    [jobId, trackerType, new Date().toISOString(), updatedRows]
  ]);
  sheet.setVisibility(ExcelScript.SheetVisibility.hidden);
}
