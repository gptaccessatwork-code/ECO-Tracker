interface CrfForwardDetectionResult {
  found: boolean;
  shouldForward: boolean;
  crfNumber: string;
  reason: string;
}

/**
 * Detects whether an Outlook message should enter the CRF forwarding flow.
 *
 * This mirrors the Outlook VBA rules:
 * - Subject must contain CRF followed by an optional :, #, or - and at least
 *   five digits.
 * - Do not forward if the Austin Change Coordinator is already a recipient.
 *
 * The workbook argument is required by Office Scripts but this script does
 * not read or modify the workbook.
 */
function main(
  workbook: ExcelScript.Workbook,
  subject: string,
  recipientsText: string
): CrfForwardDetectionResult {
  const normalizedSubject: string = (subject ?? "").trim();
  const normalizedRecipients: string = (recipientsText ?? "").toLowerCase();

  const match: RegExpMatchArray | null =
    normalizedSubject.match(/CRF\s*[:#-]?\s*(\d{5,})/i);

  if (match === null) {
    return {
      found: false,
      shouldForward: false,
      crfNumber: "",
      reason: "Subject does not contain a CRF number with at least five digits."
    };
  }

  const crfNumber: string = match[1];
  const coordinatorAlreadyIncluded: boolean =
    normalizedRecipients.includes("change_cordinator@ichorsystems.com") ||
    normalizedRecipients.includes("austin change coordinator");

  if (coordinatorAlreadyIncluded) {
    return {
      found: true,
      shouldForward: false,
      crfNumber,
      reason: "Austin Change Coordinator is already a recipient."
    };
  }

  return {
    found: true,
    shouldForward: true,
    crfNumber,
    reason: "CRF detected and Change Coordinator is not already a recipient."
  };
}
