"""Fetch ECO submitted/released dates from Agile Table WSDL.

This script connects to Agile PLM through the Table SOAP service, loads the
workflow-related table for a given ECO, and looks for workflow statuses named
"Submitted" and "Released". For each matching status, it returns the
"local client time" if present.

Credentials are read from environment variables:
- AGILE_USER
- AGILE_PASS
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional
from zoneinfo import ZoneInfo

from lxml import etree
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from requests import Session
from requests.auth import HTTPBasicAuth
from zeep import Client, Settings, helpers
from zeep.transports import Transport


AGILE_WSDL = "http://pagapps1.ichorsystems.com:7001/CoreService/services/Table?wsdl"
REQUEST_TABLE_QNAME = "{http://xmlns.oracle.com/AgileObjects/Core/Table/V1}RequestTableType"
LOAD_TABLE_REQUEST_QNAME = "{http://xmlns.oracle.com/AgileObjects/Core/Table/V1}LoadTableRequestType"
PORTLAND_TZ = ZoneInfo("America/Los_Angeles")
SINGAPORE_TZ = ZoneInfo("Asia/Singapore")

DEFAULT_CLASS_CANDIDATES = [
    "Change",
    "ECO",
    "Change Orders",
]

DEFAULT_TABLE_CANDIDATES = [
    "History",
    "Workflow",
    "Workflow Status",
]

STATUS_FIELD_CANDIDATES = [
    "status",
    "nextstatus",
    "workflowstatus",
    "currentstatus",
    "name",
]

TIME_FIELD_CANDIDATES = [
    "localclienttime",
    "clienttime",
    "time",
    "date",
    "modifieddate",
]

ACTION_FIELD_CANDIDATES = [
    "action",
]

DEFAULT_WORKBOOK_FILE = os.getenv("TARGET_WORKBOOK_FILE")
DEFAULT_WORKSHEET_NAME = os.getenv("TARGET_WORKSHEET")
DEFAULT_ECO_COLUMN = "C"
DEFAULT_SUBMITTED_COLUMN = "D"
DEFAULT_SUBMITTED_DELTA_COLUMN = "E"
DEFAULT_SUBMITTED_DELTA_EXCL_WEEKEND_COLUMN = "F"
DEFAULT_RELEASED_COLUMN = "G"
DEFAULT_RELEASED_DELTA_COLUMN = "H"
DEFAULT_RELEASED_DELTA_EXCL_WEEKEND_COLUMN = "I"
ALERT_FILL = PatternFill(fill_type="solid", start_color="FF5B5B", end_color="FF5B5B")
ALERT_FONT = Font(color="000000", bold=True)
CLEAR_FILL = PatternFill(fill_type=None)
CLEAR_FONT = Font(color="000000", bold=False)
CENTER_ALIGNMENT = Alignment(horizontal="center", vertical="center")


@dataclass
class EcoDates:
    submitted_date: Optional[str]
    released_date: Optional[str]
    class_identifier: str
    table_identifier: str


class AgileEcoClient:
    def __init__(self, username: str, password: str) -> None:
        session = Session()
        session.trust_env = False
        session.auth = HTTPBasicAuth(username, password)
        transport = Transport(session=session)
        settings = Settings(strict=False, xml_huge_tree=True)

        self.client = Client(AGILE_WSDL, transport=transport, settings=settings)
        self.RequestTableType = self.client.get_type(REQUEST_TABLE_QNAME)
        self.LoadTableRequestType = self.client.get_type(LOAD_TABLE_REQUEST_QNAME)

    def load_table(
        self,
        eco_number: str,
        class_identifier: str,
        table_identifier: str,
    ) -> List[Dict[str, Any]]:
        table_request = self.RequestTableType(
            classIdentifier=class_identifier,
            objectNumber=eco_number,
            tableIdentifier=table_identifier,
        )
        load_request = self.LoadTableRequestType(tableRequest=[table_request])
        response = self.client.service.loadTable(load_request)
        payload = helpers.serialize_object(response)
        table_contents = payload.get("tableContents") or []
        if not table_contents:
            return []
        return table_contents[0].get("row", []) or []


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Fetch Submitted/Released dates for ECOs from Agile."
    )
    parser.add_argument(
        "eco_number",
        nargs="?",
        help="Single ECO number to inspect. Omit this to update a workbook.",
    )
    parser.add_argument(
        "--workbook-file",
        default=DEFAULT_WORKBOOK_FILE,
        help="Local synced workbook file to update in place.",
    )
    parser.add_argument(
        "--worksheet",
        default=DEFAULT_WORKSHEET_NAME,
        help="Worksheet name to update. Defaults to the active sheet.",
    )
    parser.add_argument(
        "--eco-column",
        default=DEFAULT_ECO_COLUMN,
        help="Column containing ECO numbers. Default: C",
    )
    parser.add_argument(
        "--submitted-column",
        default=DEFAULT_SUBMITTED_COLUMN,
        help="Column to write Submitted Date. Default: D",
    )
    parser.add_argument(
        "--released-column",
        default=DEFAULT_RELEASED_COLUMN,
        help="Column to write Released Date. Default: G",
    )
    parser.add_argument(
        "--spec-award-column",
        default="B",
        help="Column containing Spec Award Date. Default: B",
    )
    parser.add_argument(
        "--submitted-delta-column",
        default=DEFAULT_SUBMITTED_DELTA_COLUMN,
        help="Column to write Delta (Spec to Submitted). Default: E",
    )
    parser.add_argument(
        "--submitted-delta-excl-weekend-column",
        default=DEFAULT_SUBMITTED_DELTA_EXCL_WEEKEND_COLUMN,
        help="Column to write Delta (excl weekend). Default: F",
    )
    parser.add_argument(
        "--released-delta-column",
        default=DEFAULT_RELEASED_DELTA_COLUMN,
        help="Column to write Delta (Spec to Released). Default: H",
    )
    parser.add_argument(
        "--released-delta-excl-weekend-column",
        default=DEFAULT_RELEASED_DELTA_EXCL_WEEKEND_COLUMN,
        help="Column to write Delta (excl weekend)2. Default: I",
    )
    parser.add_argument(
        "--start-row",
        type=int,
        default=2,
        help="First data row to process. Default: 2",
    )
    parser.add_argument(
        "--class-identifier",
        action="append",
        dest="class_identifiers",
        help="Override Agile class identifier. Can be used multiple times.",
    )
    parser.add_argument(
        "--table-identifier",
        action="append",
        dest="table_identifiers",
        help="Override Agile table identifier. Can be used multiple times.",
    )
    parser.add_argument(
        "--inspect",
        action="store_true",
        help="Print discovered row fields to help tune class/table identifiers.",
    )
    return parser.parse_args()


def normalize_key(name: str) -> str:
    return "".join(ch.lower() for ch in name if ch.isalnum())


def collapse_text(value: str) -> str:
    return " ".join(value.replace("\xa0", " ").split())


def extract_row_fields(row: Dict[str, Any]) -> Dict[str, str]:
    fields: Dict[str, str] = {}
    for elem in row.get("_value_1", []) or []:
        tag_name = normalize_key(etree.QName(elem.tag).localname)
        text_value = collapse_text("".join(elem.itertext()).strip())
        if text_value:
            fields[tag_name] = text_value
    return fields


def find_first_value(fields: Dict[str, str], candidates: Iterable[str]) -> Optional[str]:
    for candidate in candidates:
        key = normalize_key(candidate)
        if key in fields and fields[key]:
            return fields[key]
    return None


def parse_datetime(value: str) -> Optional[datetime]:
    formats = [
        "%Y-%m-%dT%H:%M:%S.%fZ",
        "%d %b %Y %I:%M:%S %p",
        "%d %b %Y %H:%M:%S",
        "%d %b %Y %I:%M %p",
        "%d %b %Y %H:%M",
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def convert_agile_time_to_singapore(value: str) -> Optional[datetime]:
    parsed = parse_datetime(value)
    if parsed is None:
        return None

    # Agile's "local client time" is business-local Portland time for this workflow,
    # even when the value is serialized with a trailing "Z".
    if parsed.tzinfo is not None:
        parsed = parsed.replace(tzinfo=None)

    portland_time = parsed.replace(tzinfo=PORTLAND_TZ)
    return portland_time.astimezone(SINGAPORE_TZ)


def contains_status_name(value: str, status_name: str) -> bool:
    normalized_value = value.strip().lower()
    normalized_status = status_name.strip().lower()
    return normalized_status in normalized_value


def format_output_date(value: str) -> str:
    parsed = convert_agile_time_to_singapore(value)
    if parsed is None:
        return value
    return parsed.strftime("%d %b %Y")


def parse_sheet_date(value: Any) -> Optional[date]:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text_value = str(value).strip()
    if not text_value:
        return None
    for fmt in ("%d %b %Y", "%d %B %Y", "%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text_value, fmt).date()
        except ValueError:
            continue
    return None


def calculate_delta_days(start_date: Optional[date], end_date: Optional[date]) -> Optional[int]:
    if not start_date or not end_date:
        return None
    return abs((end_date - start_date).days)


def count_weekend_days_between(start_date: date, end_date: date) -> int:
    if end_date == start_date:
        return 0
    if end_date < start_date:
        start_date, end_date = end_date, start_date
    weekend_days = 0
    current_date = start_date + timedelta(days=1)
    while current_date <= end_date:
        if current_date.weekday() >= 5:
            weekend_days += 1
        current_date += timedelta(days=1)
    return weekend_days


def calculate_delta_excluding_weekends(
    start_date: Optional[date],
    end_date: Optional[date],
) -> Optional[int]:
    if not start_date or not end_date:
        return None
    raw_delta = abs((end_date - start_date).days)
    weekend_days = count_weekend_days_between(start_date, end_date)
    return raw_delta - weekend_days


def singapore_today() -> date:
    return datetime.now(SINGAPORE_TZ).date()


def clear_cell_style(cell) -> None:
    cell.fill = CLEAR_FILL
    cell.font = CLEAR_FONT
    cell.alignment = CENTER_ALIGNMENT


def apply_conditional_formatting(
    worksheet,
    start_row: int,
    end_row: int,
    column_letter: str,
    threshold: int,
) -> None:
    if end_row < start_row:
        return

    range_ref = f"{column_letter}{start_row}:{column_letter}{end_row}"
    formula = f"AND(ISNUMBER({column_letter}{start_row}),{column_letter}{start_row}>{threshold})"
    worksheet.conditional_formatting.add(
        range_ref,
        FormulaRule(formula=[formula], fill=ALERT_FILL, font=ALERT_FONT),
    )


def clear_existing_conditional_formatting_for_columns(
    worksheet,
    start_row: int,
    end_row: int,
    columns: List[str],
) -> None:
    target_ranges = {f"{col}{start_row}:{col}{end_row}" for col in columns}
    cf_rules = getattr(worksheet.conditional_formatting, "_cf_rules", None)
    if cf_rules is None:
        return

    to_remove = []
    for key in list(cf_rules.keys()):
        sqref_text = str(getattr(key, "sqref", key))
        if sqref_text in target_ranges:
            to_remove.append(key)

    for key in to_remove:
        del cf_rules[key]


def extract_status_dates(
    rows: List[Dict[str, Any]],
    inspect: bool = False,
) -> tuple[Optional[str], Optional[str]]:
    submitted_matches: List[tuple[int, str]] = []
    released_matches: List[tuple[int, str]] = []

    for index, row in enumerate(rows, start=1):
        fields = extract_row_fields(row)
        if inspect:
            print(f"Row {index} fields: {fields}")

        action_value = find_first_value(fields, ACTION_FIELD_CANDIDATES) or ""
        time_value = find_first_value(fields, TIME_FIELD_CANDIDATES)
        if not time_value:
            continue

        status_candidates = []
        for candidate in STATUS_FIELD_CANDIDATES:
            field_key = normalize_key(candidate)
            if field_key in fields and fields[field_key]:
                status_candidates.append(fields[field_key])

        if not status_candidates:
            continue

        row_weight = 0
        if "changestatus" in normalize_key(action_value):
            row_weight = 1

        if any(contains_status_name(value, "submitted") for value in status_candidates):
            submitted_matches.append((row_weight, time_value))
        if any(contains_status_name(value, "released") for value in status_candidates):
            released_matches.append((row_weight, time_value))

    submitted_date = pick_best_date(submitted_matches)
    released_date = pick_best_date(released_matches)
    return submitted_date, released_date


def pick_best_date(values: List[tuple[int, str]]) -> Optional[str]:
    if not values:
        return None

    parsed_values = [(weight, parse_datetime(value), value) for weight, value in values]
    parsed_only = [pair for pair in parsed_values if pair[1] is not None]
    if parsed_only:
        parsed_only.sort(key=lambda pair: (-pair[0], pair[1]))
        return format_output_date(parsed_only[0][2])
    return values[0][1]


def fetch_eco_dates(
    eco_number: str,
    class_identifiers: List[str],
    table_identifiers: List[str],
    inspect: bool,
) -> EcoDates:
    agile_user = os.getenv("AGILE_USER")
    agile_pass = os.getenv("AGILE_PASS")
    if not agile_user or not agile_pass:
        raise RuntimeError("AGILE_USER and AGILE_PASS must be set in the environment.")

    client = AgileEcoClient(agile_user, agile_pass)
    errors: List[str] = []

    for class_identifier in class_identifiers:
        for table_identifier in table_identifiers:
            try:
                rows = client.load_table(eco_number, class_identifier, table_identifier)
            except Exception as exc:
                errors.append(f"{class_identifier}/{table_identifier}: {exc}")
                continue

            if not rows:
                continue

            submitted_date, released_date = extract_status_dates(rows, inspect=inspect)
            if submitted_date or released_date:
                return EcoDates(
                    submitted_date=submitted_date,
                    released_date=released_date,
                    class_identifier=class_identifier,
                    table_identifier=table_identifier,
                )

            if inspect:
                return EcoDates(
                    submitted_date=None,
                    released_date=None,
                    class_identifier=class_identifier,
                    table_identifier=table_identifier,
                )

    error_text = "\n".join(errors[-5:]) if errors else "No matching workflow rows found."
    raise RuntimeError(
        "Could not find Submitted/Released workflow dates for "
        f"{eco_number}.\nTried classes={class_identifiers} tables={table_identifiers}\n{error_text}"
    )


def update_workbook_dates(
    workbook_file: str,
    worksheet_name: Optional[str],
    spec_award_column: str,
    eco_column: str,
    submitted_column: str,
    submitted_delta_column: str,
    submitted_delta_excl_weekend_column: str,
    released_column: str,
    released_delta_column: str,
    released_delta_excl_weekend_column: str,
    start_row: int,
    class_identifiers: List[str],
    table_identifiers: List[str],
) -> int:
    workbook_path = Path(workbook_file)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook file not found: {workbook_file}")

    workbook = load_workbook(workbook_path)
    worksheet = workbook[worksheet_name] if worksheet_name else workbook.active

    spec_award_col = spec_award_column.upper()
    eco_col = eco_column.upper()
    submitted_col = submitted_column.upper()
    submitted_delta_col = submitted_delta_column.upper()
    submitted_delta_excl_col = submitted_delta_excl_weekend_column.upper()
    released_col = released_column.upper()
    released_delta_col = released_delta_column.upper()
    released_delta_excl_col = released_delta_excl_weekend_column.upper()

    updated_rows = 0
    eco_cache: Dict[str, EcoDates] = {}
    today_sg = singapore_today()
    for row_idx in range(start_row, worksheet.max_row + 1):
        eco_value = worksheet[f"{eco_col}{row_idx}"].value
        spec_award_date = parse_sheet_date(worksheet[f"{spec_award_col}{row_idx}"].value)
        eco_number = "" if eco_value is None else str(eco_value).strip()

        result: Optional[EcoDates] = None
        if eco_number:
            if eco_number not in eco_cache:
                eco_cache[eco_number] = fetch_eco_dates(
                    eco_number=eco_number,
                    class_identifiers=class_identifiers,
                    table_identifiers=table_identifiers,
                    inspect=False,
                )
            result = eco_cache[eco_number]

        submitted_date_text = result.submitted_date if result else ""
        released_date_text = result.released_date if result else ""
        worksheet[f"{submitted_col}{row_idx}"] = submitted_date_text
        worksheet[f"{released_col}{row_idx}"] = released_date_text

        submitted_date = parse_sheet_date(submitted_date_text)
        released_date = parse_sheet_date(released_date_text)
        submitted_delta_end = submitted_date or today_sg
        released_delta_end = released_date or today_sg

        submitted_delta_value = calculate_delta_days(spec_award_date, submitted_delta_end)
        submitted_delta_excl_value = calculate_delta_excluding_weekends(
            spec_award_date, submitted_delta_end
        )
        released_delta_value = calculate_delta_days(spec_award_date, released_delta_end)
        released_delta_excl_value = calculate_delta_excluding_weekends(
            spec_award_date, released_delta_end
        )

        submitted_delta_cell = worksheet[f"{submitted_delta_col}{row_idx}"]
        submitted_delta_excl_cell = worksheet[f"{submitted_delta_excl_col}{row_idx}"]
        released_delta_cell = worksheet[f"{released_delta_col}{row_idx}"]
        released_delta_excl_cell = worksheet[f"{released_delta_excl_col}{row_idx}"]

        submitted_delta_cell.value = submitted_delta_value
        submitted_delta_excl_cell.value = submitted_delta_excl_value
        released_delta_cell.value = released_delta_value
        released_delta_excl_cell.value = released_delta_excl_value

        worksheet[f"{submitted_col}{row_idx}"].alignment = CENTER_ALIGNMENT
        worksheet[f"{released_col}{row_idx}"].alignment = CENTER_ALIGNMENT

        clear_cell_style(submitted_delta_cell)
        clear_cell_style(submitted_delta_excl_cell)
        clear_cell_style(released_delta_cell)
        clear_cell_style(released_delta_excl_cell)
        updated_rows += 1
        print(
            f"Updated row {row_idx}: ECO={eco_number} "
            f"Submitted={submitted_date_text} Released={released_date_text}"
        )

    clear_existing_conditional_formatting_for_columns(
        worksheet,
        start_row,
        worksheet.max_row,
        [
            submitted_delta_col,
            released_delta_col,
            submitted_delta_excl_col,
            released_delta_excl_col,
        ],
    )
    apply_conditional_formatting(worksheet, start_row, worksheet.max_row, submitted_delta_col, 2)
    apply_conditional_formatting(worksheet, start_row, worksheet.max_row, released_delta_col, 4)

    workbook.save(workbook_path)
    workbook.close()
    print(f"Workbook updated: {workbook_file}")
    return updated_rows


def main() -> int:
    args = parse_args()
    class_identifiers = args.class_identifiers or DEFAULT_CLASS_CANDIDATES
    table_identifiers = args.table_identifiers or DEFAULT_TABLE_CANDIDATES

    if args.eco_number:
        result = fetch_eco_dates(
            eco_number=args.eco_number,
            class_identifiers=class_identifiers,
            table_identifiers=table_identifiers,
            inspect=args.inspect,
        )

        print(f"ECO: {args.eco_number}")
        print(f"Class Identifier: {result.class_identifier}")
        print(f"Table Identifier: {result.table_identifier}")
        print(f"Submitted Date: {result.submitted_date or ''}")
        print(f"Released Date: {result.released_date or ''}")
        return 0

    if not args.workbook_file:
        raise RuntimeError("Provide either an ECO number or --workbook-file.")

    update_workbook_dates(
        workbook_file=args.workbook_file,
        worksheet_name=args.worksheet,
        spec_award_column=args.spec_award_column,
        eco_column=args.eco_column,
        submitted_column=args.submitted_column,
        submitted_delta_column=args.submitted_delta_column,
        submitted_delta_excl_weekend_column=args.submitted_delta_excl_weekend_column,
        released_column=args.released_column,
        released_delta_column=args.released_delta_column,
        released_delta_excl_weekend_column=args.released_delta_excl_weekend_column,
        start_row=args.start_row,
        class_identifiers=class_identifiers,
        table_identifiers=table_identifiers,
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
