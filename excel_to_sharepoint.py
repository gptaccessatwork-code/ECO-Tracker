"""Append rows from a local Excel file into a workbook table.

Supported target modes:
- Microsoft Graph for a SharePoint-hosted workbook
- A locally synced OneDrive/SharePoint workbook file
"""

from __future__ import annotations

import argparse
import base64
import os
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional
from urllib.parse import quote

import msal
import requests
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["Files.ReadWrite"]
DEFAULT_CACHE_FILE = ".msal_token_cache.bin"
DEFAULT_BATCH_SIZE = 50


@dataclass
class Config:
    tenant_id: Optional[str]
    client_id: Optional[str]
    workbook_url: Optional[str]
    sharepoint_hostname: Optional[str]
    sharepoint_site_path: Optional[str]
    workbook_path: Optional[str]
    target_workbook_file: Optional[str]
    table_name: str
    excel_file: str
    excel_sheet: Optional[str]
    field_map: Dict[str, str]
    batch_size: int
    cache_file: str
    dry_run: bool


def parse_args() -> Config:
    parser = argparse.ArgumentParser(
        description="Append local Excel rows to a SharePoint workbook table."
    )
    parser.add_argument("--tenant-id", default=os.getenv("TENANT_ID"), required=False)
    parser.add_argument("--client-id", default=os.getenv("CLIENT_ID"), required=False)
    parser.add_argument(
        "--workbook-url",
        default=os.getenv("WORKBOOK_URL"),
        required=False,
        help="Full SharePoint/OneDrive workbook URL. Use this instead of hostname, site path, and workbook path.",
    )
    parser.add_argument(
        "--sharepoint-hostname",
        default=os.getenv("SHAREPOINT_HOSTNAME"),
        required=False,
        help="Example: contoso.sharepoint.com",
    )
    parser.add_argument(
        "--sharepoint-site-path",
        default=os.getenv("SHAREPOINT_SITE_PATH"),
        required=False,
        help="Example: /sites/ECO-Tracker",
    )
    parser.add_argument(
        "--workbook-path",
        default=os.getenv("WORKBOOK_PATH"),
        required=False,
        help="Path to the workbook inside the SharePoint document library, for example Shared Documents/Tracker.xlsx",
    )
    parser.add_argument(
        "--target-workbook-file",
        default=os.getenv("TARGET_WORKBOOK_FILE"),
        required=False,
        help="Local path to a synced OneDrive/SharePoint workbook file. Use this to avoid Graph authentication.",
    )
    parser.add_argument(
        "--table-name",
        default=os.getenv("TABLE_NAME"),
        required=False,
        help="Existing Excel table name in the SharePoint workbook",
    )
    parser.add_argument("--excel-file", default=os.getenv("EXCEL_FILE"), required=False)
    parser.add_argument("--excel-sheet", default=os.getenv("EXCEL_SHEET"), required=False)
    parser.add_argument(
        "--map",
        action="append",
        default=[],
        metavar="EXCEL_HEADER=TABLE_COLUMN",
        help="Map an Excel header to a workbook table column name. Can be used multiple times.",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=int(os.getenv("BATCH_SIZE", str(DEFAULT_BATCH_SIZE))),
        help="How many rows to send in one Graph request.",
    )
    parser.add_argument(
        "--cache-file",
        default=os.getenv("TOKEN_CACHE_FILE", DEFAULT_CACHE_FILE),
        help="Path to the local MSAL token cache file.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        default=os.getenv("DRY_RUN", "false").lower() in {"1", "true", "yes"},
        help="Print rows that would be uploaded without writing to SharePoint.",
    )

    args = parser.parse_args()

    missing = [
        name
        for name, value in [
            ("table_name", args.table_name),
            ("excel_file", args.excel_file),
        ]
        if not value
    ]
    use_local_target = bool(args.target_workbook_file)
    if not use_local_target and not args.tenant_id:
        missing.append("tenant_id")
    if not use_local_target and not args.client_id:
        missing.append("client_id")
    if not use_local_target and not args.workbook_url:
        missing.extend(
            name
            for name, value in [
                ("sharepoint_hostname", args.sharepoint_hostname),
                ("sharepoint_site_path", args.sharepoint_site_path),
                ("workbook_path", args.workbook_path),
            ]
            if not value
        )
    if missing:
        parser.error("Missing required values: " + ", ".join(missing))

    if args.batch_size < 1:
        parser.error("--batch-size must be at least 1")

    field_map: Dict[str, str] = {}
    for mapping in args.map:
        if "=" not in mapping:
            parser.error(f"Invalid --map value '{mapping}'. Use EXCEL_HEADER=TABLE_COLUMN.")
        excel_header, table_column = mapping.split("=", 1)
        field_map[excel_header.strip()] = table_column.strip()

    return Config(
        tenant_id=args.tenant_id,
        client_id=args.client_id,
        workbook_url=args.workbook_url,
        sharepoint_hostname=args.sharepoint_hostname,
        sharepoint_site_path=args.sharepoint_site_path,
        workbook_path=args.workbook_path,
        target_workbook_file=args.target_workbook_file,
        table_name=args.table_name,
        excel_file=args.excel_file,
        excel_sheet=args.excel_sheet,
        field_map=field_map,
        batch_size=args.batch_size,
        cache_file=args.cache_file,
        dry_run=args.dry_run,
    )


def load_cache(path: str) -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    cache_path = Path(path)
    if cache_path.exists():
        cache.deserialize(cache_path.read_text(encoding="utf-8"))
    return cache


def save_cache(cache: msal.SerializableTokenCache, path: str) -> None:
    if not cache.has_state_changed:
        return
    Path(path).write_text(cache.serialize(), encoding="utf-8")


def get_access_token(config: Config) -> str:
    if not config.tenant_id or not config.client_id:
        raise ValueError("Tenant ID and client ID are required for Graph mode.")
    cache = load_cache(config.cache_file)
    app = msal.PublicClientApplication(
        client_id=config.client_id,
        authority=f"https://login.microsoftonline.com/{config.tenant_id}",
        token_cache=cache,
    )

    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
        if "user_code" not in flow:
            raise RuntimeError("Could not start device code flow.")
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    save_cache(cache, config.cache_file)

    if not result or "access_token" not in result:
        error_text = "unknown error"
        if result:
            error_text = result.get("error_description") or result.get("error") or error_text
        raise RuntimeError("Could not acquire Graph access token: " + error_text)

    return result["access_token"]


def graph_request(method: str, url: str, token: str, **kwargs: Any) -> requests.Response:
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"
    headers.setdefault("Content-Type", "application/json")
    response = requests.request(method, url, headers=headers, timeout=60, **kwargs)
    response.raise_for_status()
    return response


def normalize_site_path(path: str) -> str:
    path = path.strip()
    if not path.startswith("/"):
        path = "/" + path
    return path


def encode_workbook_path(path: str) -> str:
    return quote(path.strip("/"), safe="/")


def encode_sharing_url(url: str) -> str:
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8")
    return "u!" + encoded.rstrip("=")


def resolve_site_id(config: Config, token: str) -> str:
    if not config.sharepoint_hostname or not config.sharepoint_site_path:
        raise ValueError("SharePoint hostname and site path are required without workbook URL.")
    site_path = normalize_site_path(config.sharepoint_site_path)
    url = f"{GRAPH_BASE_URL}/sites/{config.sharepoint_hostname}:{site_path}"
    response = graph_request("GET", url, token)
    return response.json()["id"]


def resolve_workbook_by_path_url(site_id: str, config: Config) -> str:
    if not config.workbook_path:
        raise ValueError("Workbook path is required without workbook URL.")
    workbook_path = encode_workbook_path(config.workbook_path)
    return (
        f"{GRAPH_BASE_URL}/sites/{site_id}/drive/root:/{workbook_path}:"
        f"/workbook"
    )


def resolve_workbook_by_sharing_url(config: Config, token: str) -> str:
    if not config.workbook_url:
        raise ValueError("Workbook URL is required.")

    share_id = encode_sharing_url(config.workbook_url)
    url = f"{GRAPH_BASE_URL}/shares/{share_id}/driveItem"
    response = graph_request("GET", url, token)
    drive_item = response.json()
    drive_id = drive_item.get("parentReference", {}).get("driveId")
    item_id = drive_item.get("id")

    if not drive_id or not item_id:
        raise RuntimeError("Could not resolve workbook drive and item IDs from the URL.")

    return f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{item_id}/workbook"


def resolve_workbook_base_url(config: Config, token: str) -> str:
    if config.workbook_url:
        return resolve_workbook_by_sharing_url(config, token)

    site_id = resolve_site_id(config, token)
    return resolve_workbook_by_path_url(site_id, config)


def get_table_columns(workbook_base_url: str, table_name: str, token: str) -> List[str]:
    url = f"{workbook_base_url}/tables/{quote(table_name)}/columns?$select=name"
    response = graph_request("GET", url, token)
    payload = response.json()
    columns = [item["name"] for item in payload.get("value", []) if item.get("name")]
    if not columns:
        raise RuntimeError(
            f"Could not read columns for table '{table_name}'. Make sure the table exists."
        )
    return columns


def get_local_table(target_workbook_file: str, table_name: str):
    workbook = load_workbook(target_workbook_file)
    for worksheet in workbook.worksheets:
        if table_name in worksheet.tables:
            return workbook, worksheet, worksheet.tables[table_name]
        for table in worksheet.tables.values():
            if getattr(table, "displayName", None) == table_name:
                return workbook, worksheet, table
    workbook.close()
    raise RuntimeError(
        f"Could not find table '{table_name}' in local workbook '{target_workbook_file}'."
    )


def get_local_table_columns(worksheet, table) -> List[str]:
    min_col, min_row, max_col, _ = range_boundaries(table.ref)
    headers: List[str] = []
    for col_idx in range(min_col, max_col + 1):
        value = worksheet.cell(row=min_row, column=col_idx).value
        headers.append("" if value is None else str(value).strip())
    return headers


def read_local_table_columns(target_workbook_file: str, table_name: str) -> List[str]:
    workbook, worksheet, table = get_local_table(target_workbook_file, table_name)
    try:
        return get_local_table_columns(worksheet, table)
    finally:
        workbook.close()


def append_rows_to_local_table(
    target_workbook_file: str,
    table_name: str,
    rows: List[List[Any]],
) -> None:
    print(f"Opening target workbook: {target_workbook_file}")
    workbook, worksheet, table = get_local_table(target_workbook_file, table_name)
    if getattr(table, "totalsRowShown", False):
        workbook.close()
        raise RuntimeError(
            f"Table '{table_name}' uses a totals row. This script currently expects a normal table without totals."
        )

    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    next_row = max_row + 1

    print(f"Writing {len(rows)} row(s) into table '{table_name}'...")
    for row_values in rows:
        for offset, value in enumerate(row_values):
            worksheet.cell(row=next_row, column=min_col + offset, value=value)
        next_row += 1

    new_max_row = max_row + len(rows)
    table.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{new_max_row}"
    print("Saving workbook to local synced file...")
    workbook.save(target_workbook_file)
    workbook.close()
    print("Local workbook save complete.")


def load_rows_from_excel(path: str, sheet_name: Optional[str]) -> List[Dict[str, Any]]:
    workbook = load_workbook(path, data_only=True)
    sheet = workbook[sheet_name] if sheet_name else workbook.active

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []

    headers = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
    data_rows: List[Dict[str, Any]] = []

    for row in rows[1:]:
        item: Dict[str, Any] = {}
        for header, value in zip(headers, row):
            if not header or value is None:
                continue
            item[header] = value
        if item:
            data_rows.append(item)

    return data_rows


def normalize_value(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float, str)):
        return value
    return str(value)


def apply_field_map(row: Dict[str, Any], field_map: Dict[str, str]) -> Dict[str, Any]:
    mapped: Dict[str, Any] = {}
    for excel_header, value in row.items():
        table_column = field_map.get(excel_header, excel_header)
        mapped[table_column] = normalize_value(value)
    return mapped


def align_row_to_table(row: Dict[str, Any], table_columns: List[str]) -> List[Any]:
    return [row.get(column) for column in table_columns]


def batch_rows(items: List[List[Any]], batch_size: int) -> Iterable[List[List[Any]]]:
    for start in range(0, len(items), batch_size):
        yield items[start : start + batch_size]


def append_rows_to_table(
    workbook_base_url: str,
    table_name: str,
    token: str,
    rows: List[List[Any]],
) -> None:
    url = f"{workbook_base_url}/tables/{quote(table_name)}/rows/add"
    payload = {"values": rows}
    graph_request("POST", url, token, json=payload)


def main() -> int:
    config = parse_args()
    source_rows = load_rows_from_excel(config.excel_file, config.excel_sheet)

    if not source_rows:
        print("No data rows found in the Excel file.")
        return 0

    workbook_base_url = None
    token = None
    if config.target_workbook_file:
        table_columns = read_local_table_columns(
            config.target_workbook_file, config.table_name
        )
    else:
        token = get_access_token(config)
        workbook_base_url = resolve_workbook_base_url(config, token)
        table_columns = get_table_columns(workbook_base_url, config.table_name, token)

    print(f"Found {len(source_rows)} source row(s).")
    print(
        f"Target workbook: "
        f"{config.target_workbook_file or config.workbook_url or config.workbook_path}"
    )
    print(f"Target table: {config.table_name}")
    print(f"Table columns: {', '.join(table_columns)}")

    prepared_rows: List[List[Any]] = []
    for index, row in enumerate(source_rows, start=2):
        mapped = apply_field_map(row, config.field_map)
        aligned = align_row_to_table(mapped, table_columns)
        prepared_rows.append(aligned)
        if config.dry_run:
            print(f"[DRY RUN] Row {index}: {aligned}")

    if config.dry_run:
        print("Dry run complete. No SharePoint data was changed.")
        return 0

    if config.target_workbook_file:
        append_rows_to_local_table(config.target_workbook_file, config.table_name, prepared_rows)
        written = len(prepared_rows)
        print(f"Appended {written}/{len(prepared_rows)} row(s).")
    else:
        written = 0
        for chunk in batch_rows(prepared_rows, config.batch_size):
            append_rows_to_table(workbook_base_url, config.table_name, token, chunk)
            written += len(chunk)
            print(f"Appended {written}/{len(prepared_rows)} row(s).")

    print(f"Upload complete. {written} row(s) appended.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except requests.HTTPError as exc:
        print(f"Graph request failed: {exc}", file=sys.stderr)
        if exc.response is not None:
            print(exc.response.text, file=sys.stderr)
        raise SystemExit(1)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
