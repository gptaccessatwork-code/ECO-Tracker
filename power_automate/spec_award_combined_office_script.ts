interface SpecAwardCombinedResult {
  detected: boolean;
  workbookWriteEnabled: boolean;
  added: number;
  skippedDuplicates: number;
  systemNumbers: string[];
  addedSystemNumbers: string[];
  forwardHtml: string;
  reason: string;
}

/**
 * Parse a Spec Award email, optionally append new systems to Table1, and
 * return the cleaned HTML that Power Automate should forward.
 *
 * Power Automate inputs:
 * - emailHtml: Body from "When a new email arrives (V3)"
 * - receivedTime: Received Time from "When a new email arrives (V3)"
 * - writeToWorkbook: false in TEST mode and true in LIVE mode
 *
 * The target table must contain "System Number" and "Spec Award Date".
 * TEST mode performs no workbook reads or writes after parsing the email.
 */
function main(
  workbook: ExcelScript.Workbook,
  emailHtml: string,
  receivedTime: string,
  writeToWorkbook: boolean
): SpecAwardCombinedResult {
  const tableName = "Table1";
  const systemNumberColumn = "System Number";
  const specAwardDateColumn = "Spec Award Date";
  const minimumForwardBodyLength = 200;

  if (!emailHtml || emailHtml.trim().length === 0) {
    throw new Error("The email HTML body is empty.");
  }

  const latestHtml = getLatestBodyOnly(emailHtml);
  const forwardHtml =
    latestHtml.length < minimumForwardBodyLength ? emailHtml : latestHtml;
  const latestText = cleanHtmlText(latestHtml).toLowerCase();
  const hasRequiredKeywords =
    latestText.includes("part number") &&
    latestText.includes("qtr") &&
    (latestText.includes("need by date") || latestText.includes("nbd"));

  if (!hasRequiredKeywords) {
    const reason =
      "Required keywords were not found in the latest email body. Expected Part Number, Qtr, and Need By Date or NBD.";
    console.log(reason);
    return {
      detected: false,
      workbookWriteEnabled: writeToWorkbook,
      added: 0,
      skippedDuplicates: 0,
      systemNumbers: [],
      addedSystemNumbers: [],
      forwardHtml: "",
      reason
    };
  }

  const parsedSystemNumbers = extractSystemNumbers(latestHtml);
  if (parsedSystemNumbers.length === 0) {
    const reason =
      "No HTML table containing Slot/Slot Number and Part Number data rows was found.";
    console.log(reason);
    return {
      detected: false,
      workbookWriteEnabled: writeToWorkbook,
      added: 0,
      skippedDuplicates: 0,
      systemNumbers: [],
      addedSystemNumbers: [],
      forwardHtml: "",
      reason
    };
  }

  const uniqueSystemNumbers: string[] = [];
  const messageSystemNumberKeys = new Set<string>();
  let withinMessageDuplicates = 0;

  parsedSystemNumbers.forEach((systemNumber) => {
    const key = normalizeKey(systemNumber);
    if (!key || messageSystemNumberKeys.has(key)) {
      withinMessageDuplicates += 1;
      return;
    }

    messageSystemNumberKeys.add(key);
    uniqueSystemNumbers.push(systemNumber);
  });

  if (!writeToWorkbook) {
    const reason =
      "Valid Spec Award detected. TEST mode is active, so the workbook was not changed.";
    console.log(reason);
    return {
      detected: true,
      workbookWriteEnabled: false,
      added: 0,
      skippedDuplicates: withinMessageDuplicates,
      systemNumbers: uniqueSystemNumbers,
      addedSystemNumbers: [],
      forwardHtml,
      reason
    };
  }

  const specAwardDate = toSingaporeDate(receivedTime);
  const workbookTables = workbook.getTables();
  let table: ExcelScript.Table | undefined;

  for (let tableIndex = 0; tableIndex < workbookTables.length; tableIndex += 1) {
    const candidate = workbookTables[tableIndex];
    if (candidate.getName().toLowerCase() === tableName.toLowerCase()) {
      table = candidate;
      break;
    }
  }

  if (!table) {
    throw new Error(`Could not find Excel table '${tableName}'.`);
  }

  const headers = table
    .getHeaderRowRange()
    .getTexts()[0]
    .map((value) => value.trim());
  const systemNumberIndex = findHeaderIndex(headers, systemNumberColumn);
  const specAwardDateIndex = findHeaderIndex(headers, specAwardDateColumn);

  if (systemNumberIndex < 0) {
    throw new Error(
      `Table '${tableName}' does not contain '${systemNumberColumn}'.`
    );
  }
  if (specAwardDateIndex < 0) {
    throw new Error(
      `Table '${tableName}' does not contain '${specAwardDateColumn}'.`
    );
  }

  const existingSystemNumbers = new Set<string>();
  if (table.getRowCount() > 0) {
    const existingValues = table
      .getColumnByName(systemNumberColumn)
      .getRangeBetweenHeaderAndTotal()
      .getTexts();

    existingValues.forEach((row) => {
      const normalized = normalizeKey(row[0]);
      if (normalized) {
        existingSystemNumbers.add(normalized);
      }
    });
  }

  const rowsToAdd: (string | number | boolean)[][] = [];
  const addedSystemNumbers: string[] = [];
  let skippedDuplicates = 0;

  parsedSystemNumbers.forEach((systemNumber) => {
    const key = normalizeKey(systemNumber);
    if (!key || existingSystemNumbers.has(key)) {
      skippedDuplicates += 1;
      return;
    }

    const newRow: (string | number | boolean)[] = headers.map(() => "");
    newRow[systemNumberIndex] = systemNumber;
    newRow[specAwardDateIndex] = toExcelDateSerial(specAwardDate);
    rowsToAdd.push(newRow);
    addedSystemNumbers.push(systemNumber);
    existingSystemNumbers.add(key);
  });

  if (rowsToAdd.length > 0) {
    const protection = table.getWorksheet().getProtection();
    let protectionPaused = false;

    try {
      if (protection.getProtected() && !protection.getIsPaused()) {
        if (!protection.getCanPauseProtection()) {
          throw new Error(
            "The target worksheet is protected and its protection cannot be paused."
          );
        }

        if (protection.getIsPasswordProtected()) {
          throw new Error(
            "The target worksheet has password protection. This flow is configured for passwordless sheet protection."
          );
        }

        protection.pauseProtection();
        protectionPaused = true;
      }

      const firstAddedRowIndex = table.getRowCount();
      table.addRows(-1, rowsToAdd);

      const addedDateRange = table
        .getColumnByName(specAwardDateColumn)
        .getRangeBetweenHeaderAndTotal()
        .getCell(firstAddedRowIndex, 0)
        .getResizedRange(rowsToAdd.length - 1, 0);
      addedDateRange.setNumberFormat("dd mmm yyyy");
      addedDateRange
        .getFormat()
        .setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    } finally {
      if (protectionPaused) {
        protection.resumeProtection();
      }
    }
  }

  const reason =
    rowsToAdd.length > 0
      ? "Valid Spec Award detected and new System Number rows were added."
      : "Valid Spec Award detected; all parsed System Numbers already existed.";
  console.log(reason);

  return {
    detected: true,
    workbookWriteEnabled: true,
    added: rowsToAdd.length,
    skippedDuplicates,
    systemNumbers: uniqueSystemNumbers,
    addedSystemNumbers,
    forwardHtml,
    reason
  };
}

function findHeaderIndex(headers: string[], expectedHeader: string): number {
  const expected = expectedHeader.trim().toLowerCase();
  return headers.findIndex((header) => header.trim().toLowerCase() === expected);
}

function toSingaporeDate(receivedTime: string): string {
  const timestamp = Date.parse(receivedTime);
  if (isNaN(timestamp)) {
    throw new Error(`Received Time is not a valid date/time: ${receivedTime}`);
  }

  const singaporeTime = new Date(timestamp + 8 * 60 * 60 * 1000);
  const year = singaporeTime.getUTCFullYear();
  const month = String(singaporeTime.getUTCMonth() + 1).padStart(2, "0");
  const day = String(singaporeTime.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function toExcelDateSerial(isoDate: string): number {
  const parts = isoDate.split("-");
  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);
  return Date.UTC(year, month - 1, day) / 86400000 + 25569;
}

function normalizeKey(value: string): string {
  return collapseWhitespace(value).toUpperCase();
}

function getLatestBodyOnly(html: string): string {
  // Mimecast CyberGraph inserts hidden preview rows into the message HTML.
  // They are not visible to the user and must not participate in detection.
  const visibleHtml = removeNonVisibleHtml(html);

  const replyMarker = /<div\b[^>]*\bid\s*=\s*["']?divrplyfwdmsg["']?[^>]*>/i;
  const replyMatch = replyMarker.exec(visibleHtml);
  if (replyMatch && replyMatch.index > 0) {
    return visibleHtml.slice(0, replyMatch.index);
  }

  // Outlook/Word replies frequently have no divRplyFwdMsg. Instead, quoted
  // history begins in a bordered div containing From/Sent/To/Subject labels.
  const outlookReplyIndex = findOutlookReplyHeaderIndex(visibleHtml);
  if (outlookReplyIndex > 0) {
    return visibleHtml.slice(0, outlookReplyIndex);
  }

  const horizontalRuleIndex = visibleHtml.search(/<hr\b/i);
  if (horizontalRuleIndex > 10000) {
    return visibleHtml.slice(0, horizontalRuleIndex);
  }

  return visibleHtml;
}

function removeNonVisibleHtml(html: string): string {
  return html
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(
      /<tr\b(?=[^>]*\bstyle\s*=\s*["'][^"']*display\s*:\s*none)[^>]*>[\s\S]*?<\/tr>/gi,
      ""
    );
}

function findOutlookReplyHeaderIndex(html: string): number {
  const divPattern = /<div\b[^>]*>/gi;
  let match: RegExpExecArray | null;

  while ((match = divPattern.exec(html)) !== null) {
    const openingTag = match[0];
    if (!/border-top\s*:/i.test(openingTag)) {
      continue;
    }

    // Long recipient lists can put Subject several thousand characters after
    // the opening div, so inspect a bounded section rather than one paragraph.
    const candidate = html.slice(match.index, match.index + 20000);
    const candidateText = cleanHtmlText(candidate).toLowerCase();
    if (
      candidateText.includes("from:") &&
      candidateText.includes("sent:") &&
      candidateText.includes("to:") &&
      candidateText.includes("subject:")
    ) {
      return match.index;
    }
  }

  return -1;
}

function extractSystemNumbers(html: string): string[] {
  const tables = html.match(/<table\b[\s\S]*?<\/table>/gi) ?? [];

  for (const tableHtml of tables) {
    const rows = tableHtml.match(/<tr\b[\s\S]*?<\/tr>/gi) ?? [];
    if (rows.length === 0) {
      continue;
    }

    const headerCells = extractCells(rows[0]);
    const normalizedHeaders = headerCells.map((value) => normalizeHeader(value));
    const slotNumberIndex = findFirstHeaderIndex(normalizedHeaders, [
      "slotnumber",
      "slot"
    ]);
    const partNumberIndex = normalizedHeaders.indexOf("partnumber");

    if (slotNumberIndex < 0 || partNumberIndex < 0) {
      continue;
    }

    const results: string[] = [];
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const cells = extractCells(rows[rowIndex]);
      if (
        slotNumberIndex >= cells.length ||
        partNumberIndex >= cells.length
      ) {
        continue;
      }

      const slotNumber = cleanCellText(cells[slotNumberIndex]);
      const partNumber = cleanCellText(cells[partNumberIndex]);
      const systemNumber = buildSystemNumber(slotNumber, partNumber);
      if (systemNumber) {
        results.push(systemNumber);
      }
    }

    if (results.length > 0) {
      return results;
    }
  }

  return [];
}

function findFirstHeaderIndex(
  normalizedHeaders: string[],
  acceptedHeaders: string[]
): number {
  for (let headerIndex = 0; headerIndex < normalizedHeaders.length; headerIndex += 1) {
    if (acceptedHeaders.includes(normalizedHeaders[headerIndex])) {
      return headerIndex;
    }
  }
  return -1;
}

function extractCells(rowHtml: string): string[] {
  const results: string[] = [];
  const cellPattern = /<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/gi;
  let match: RegExpExecArray | null;

  while ((match = cellPattern.exec(rowHtml)) !== null) {
    results.push(match[1]);
  }

  return results;
}

function buildSystemNumber(slotNumber: string, partNumber: string): string {
  const cleanSlot = collapseWhitespace(slotNumber);
  const cleanPart = collapseWhitespace(partNumber);
  if (!cleanSlot || !cleanPart) {
    return "";
  }

  if (/SEMSYS/i.test(cleanPart)) {
    return cleanPart.replace(/SEMSYS/i, cleanSlot);
  }
  if (/SEMNSO/i.test(cleanPart)) {
    return cleanPart.replace(/SEMNSO/i, cleanSlot);
  }
  return cleanPart;
}

function normalizeHeader(value: string): string {
  return cleanCellText(value)
    .toLowerCase()
    .replace(/[\s\-/]/g, "");
}

function cleanCellText(value: string): string {
  return collapseWhitespace(cleanHtmlText(value));
}

function cleanHtmlText(value: string): string {
  const withoutTags = value
    .replace(/<br\s*\/?>/gi, " ")
    .replace(/<\/p\s*>/gi, " ")
    .replace(/<[^>]+>/g, " ");

  return decodeHtmlEntities(withoutTags);
}

function decodeHtmlEntities(value: string): string {
  const namedEntities: { [key: string]: string } = {
    amp: "&",
    apos: "'",
    gt: ">",
    lt: "<",
    nbsp: " ",
    quot: "\""
  };

  return value.replace(
    /&(#x[0-9a-f]+|#\d+|amp|apos|gt|lt|nbsp|quot);/gi,
    (_wholeMatch: string, entity: string) => {
      const normalized = entity.toLowerCase();
      if (normalized.startsWith("#x")) {
        return String.fromCharCode(parseInt(normalized.slice(2), 16));
      }
      if (normalized.startsWith("#")) {
        return String.fromCharCode(parseInt(normalized.slice(1), 10));
      }
      return namedEntities[normalized] ?? _wholeMatch;
    }
  );
}

function collapseWhitespace(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}
