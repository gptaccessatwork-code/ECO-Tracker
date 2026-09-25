interface AgileLookupRequestItem {
  recordType: string;
  recordId: string;
  sourceRow: number;
  systemNumber: string;
  trackingNumber: string;
  baseDate: string | null;
}

/**
 * Export unreleased ECO or CRF tracker rows for the local Agile email worker.
 * The returned value is a complete JSON attachment body.
 */
function main(
  workbook: ExcelScript.Workbook,
  trackerType: string,
  jobId: string,
  tableName: string = "Table1"
): string {
  const type = trackerType.trim().toUpperCase();
  if (type !== "ECO" && type !== "CRF") {
    throw new Error("trackerType must be ECO or CRF.");
  }
  if (!jobId || !jobId.trim()) {
    throw new Error("jobId is required.");
  }

  const table = findTable(workbook, tableName);
  const headers = table.getHeaderRowRange().getTexts()[0];
  const trackingIndex = requireHeader(
    headers,
    type === "ECO" ? ["ECO Number"] : ["CRF Number"]
  );
  const baseDateIndex = requireHeader(
    headers,
    type === "ECO"
      ? ["Spec Award Date"]
      : ["CRF Date", "Received Date", "CRF Received Date"]
  );
  const releasedIndex = requireHeader(headers, ["Released Date"]);
  const systemIndex =
    type === "ECO" ? requireHeader(headers, ["System Number"]) : -1;

  const items: AgileLookupRequestItem[] = [];
  if (table.getRowCount() > 0) {
    const body = table.getRangeBetweenHeaderAndTotal();
    const values = body.getValues();
    const texts = body.getTexts();

    for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
      const releasedDate = toIsoDate(values[rowIndex][releasedIndex], texts[rowIndex][releasedIndex]);
      if (releasedDate) {
        continue;
      }

      const trackingNumber = texts[rowIndex][trackingIndex].trim();
      const systemNumber = systemIndex >= 0 ? texts[rowIndex][systemIndex].trim() : "";
      const baseDate = toIsoDate(values[rowIndex][baseDateIndex], texts[rowIndex][baseDateIndex]);
      if (!trackingNumber && !systemNumber && !baseDate) {
        continue;
      }

      // sourceRow is zero-based within the Excel table body. The identifying
      // values are also sent so the import script can detect a sort/edit and
      // find the row safely instead of updating the wrong record.
      const identity = type === "ECO" ? systemNumber : trackingNumber;
      items.push({
        recordType: type,
        recordId: `${type}:${rowIndex}:${identity || "blank"}`,
        sourceRow: rowIndex,
        systemNumber,
        trackingNumber,
        baseDate
      });
    }
  }

  return JSON.stringify({
    schemaVersion: "1.0",
    eventType: "agile.lookup.request",
    jobId: jobId.trim(),
    trackerType: type,
    createdAt: new Date().toISOString(),
    items
  });
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

function toIsoDate(value: string | number | boolean, displayedText: string): string | null {
  if (typeof value === "number" && isFinite(value) && value > 0) {
    const milliseconds = Math.round((value - 25569) * 86400000);
    const parsed = new Date(milliseconds);
    return formatUtcDate(parsed);
  }

  const text = String(displayedText || value || "").trim();
  if (!text) {
    return null;
  }
  const isoMatch = /^(\d{4})-(\d{2})-(\d{2})/.exec(text);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  }
  const parsedMilliseconds = Date.parse(text);
  if (isNaN(parsedMilliseconds)) {
    throw new Error(`Could not parse workbook date '${text}'.`);
  }
  return formatUtcDate(new Date(parsedMilliseconds));
}

function formatUtcDate(value: Date): string {
  return `${value.getUTCFullYear()}-${String(value.getUTCMonth() + 1).padStart(2, "0")}-${String(
    value.getUTCDate()
  ).padStart(2, "0")}`;
}
