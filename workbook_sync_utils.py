from __future__ import annotations

import ctypes
import os
import posixpath
import time
import traceback
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Any

WINDOWS_LOCK_ERRORS = {5, 32, 33}
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ET.register_namespace("", MAIN_NS)


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _resolve_worksheet_part(target: str) -> str:
    target = target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return posixpath.normpath(posixpath.join("xl", target))


def snapshot_worksheet_protection(workbook_path: Path | str) -> dict[str, dict[str, str]]:
    snapshot: dict[str, dict[str, str]] = {}
    with zipfile.ZipFile(workbook_path, "r") as archive:
        workbook_xml = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels_xml
            if rel.attrib.get("Id") and rel.attrib.get("Target")
        }

        for sheet in workbook_xml.findall(f"{{{MAIN_NS}}}sheets/{{{MAIN_NS}}}sheet"):
            rel_id = sheet.attrib.get(f"{{{REL_NS}}}id")
            target = rel_targets.get(rel_id or "")
            if not target:
                continue

            part_name = _resolve_worksheet_part(target)
            try:
                worksheet_xml = ET.fromstring(archive.read(part_name))
            except KeyError:
                continue

            protected_parts: dict[str, str] = {}
            for element_name in ("sheetProtection", "protectedRanges"):
                element = worksheet_xml.find(f"{{{MAIN_NS}}}{element_name}")
                if element is not None:
                    protected_parts[element_name] = ET.tostring(element, encoding="unicode")

            if protected_parts:
                snapshot[part_name] = protected_parts

    return snapshot


def _apply_protection_snapshot_to_worksheet_xml(
    worksheet_xml: bytes,
    protected_parts: dict[str, str],
) -> bytes:
    if not protected_parts:
        return worksheet_xml

    root = ET.fromstring(worksheet_xml)
    children = list(root)

    for element_name in ("sheetProtection", "protectedRanges"):
        existing = root.find(f"{{{MAIN_NS}}}{element_name}")
        if existing is not None:
            root.remove(existing)

    sheet_data_index = next(
        (index for index, child in enumerate(children) if _local_name(child.tag) == "sheetData"),
        -1,
    )
    sheet_protection_index = next(
        (index for index, child in enumerate(list(root)) if _local_name(child.tag) == "sheetProtection"),
        -1,
    )
    insert_at = sheet_protection_index + 1 if sheet_protection_index >= 0 else sheet_data_index + 1
    if insert_at <= 0:
        insert_at = len(root)

    sheet_protection_xml = protected_parts.get("sheetProtection")
    protected_ranges_xml = protected_parts.get("protectedRanges")

    if sheet_protection_xml:
        root.insert(insert_at, ET.fromstring(sheet_protection_xml))
        insert_at += 1
    if protected_ranges_xml:
        root.insert(insert_at, ET.fromstring(protected_ranges_xml))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def preserve_workbook_protection(source_path: Path | str, target_path: Path | str, log_file: Path) -> None:
    snapshot = snapshot_worksheet_protection(source_path)
    if not snapshot:
        return

    target_path = Path(target_path)
    temp_path = target_path.with_name(f"{target_path.stem}.protected{target_path.suffix}")

    try:
        with zipfile.ZipFile(target_path, "r") as source_archive, zipfile.ZipFile(
            temp_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
        ) as target_archive:
            for info in source_archive.infolist():
                data = source_archive.read(info.filename)
                protected_parts = snapshot.get(info.filename)
                if protected_parts and info.filename.startswith("xl/worksheets/"):
                    data = _apply_protection_snapshot_to_worksheet_xml(data, protected_parts)
                target_archive.writestr(info, data)

        os.replace(temp_path, target_path)
        log_event("Workbook protection and editable-range metadata preserved.", log_file)
    except Exception:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass
        raise


def is_local_conflict_error(exc: Exception) -> bool:
    if isinstance(exc, PermissionError):
        return True

    winerror = getattr(exc, "winerror", None)
    if winerror in WINDOWS_LOCK_ERRORS:
        return True

    text = str(exc).lower()
    return "being used by another process" in text or "permission denied" in text


def is_workbook_open(workbook_path: Path | str) -> bool:
    kernel32 = ctypes.windll.kernel32
    kernel32.CreateFileW.argtypes = [
        ctypes.c_wchar_p,
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.c_void_p,
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.c_void_p,
    ]
    kernel32.CreateFileW.restype = ctypes.c_void_p
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    kernel32.CloseHandle.restype = ctypes.c_int

    handle = kernel32.CreateFileW(
        str(workbook_path),
        0x80000000 | 0x40000000,
        0,
        None,
        3,
        0x80,
        None,
    )

    if handle in (None, ctypes.c_void_p(-1).value):
        return True

    kernel32.CloseHandle(handle)
    return False


def log_event(message: str, log_file: Path) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line)
    try:
        with log_file.open("a", encoding="utf-8") as handle:
            handle.write(line + "\n")
    except Exception:
        pass


def log_exception(message: str, exc: Exception, log_file: Path) -> None:
    log_event(message, log_file)
    log_event("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)).rstrip(), log_file)


def wait_for_workbook_ready(
    workbook_path: Path | str,
    retry_delay_minutes: int,
    settle_seconds: int,
    log_file: Path,
) -> None:
    while True:
        if is_workbook_open(workbook_path):
            log_event(
                f"Workbook is still open or locked. Retrying in {retry_delay_minutes} minutes.",
                log_file,
            )
            time.sleep(retry_delay_minutes * 60)
            continue

        log_event(
            f"Workbook looks free. Waiting {settle_seconds} seconds for OneDrive to settle...",
            log_file,
        )
        time.sleep(settle_seconds)

        if not is_workbook_open(workbook_path):
            return

        log_event(
            f"Workbook became busy again. Retrying in {retry_delay_minutes} minutes.",
            log_file,
        )
        time.sleep(retry_delay_minutes * 60)


def save_workbook_with_retry(
    workbook: Any,
    workbook_path: Path | str,
    retry_delay_minutes: int,
    settle_seconds: int,
    log_file: Path,
) -> None:
    while True:
        wait_for_workbook_ready(workbook_path, retry_delay_minutes, settle_seconds, log_file)
        try:
            log_event(f"Saving workbook to {workbook_path}...", log_file)
            workbook_path = Path(workbook_path)
            temp_path = Path(
                tempfile.NamedTemporaryFile(
                    mode="w+b",
                    suffix=workbook_path.suffix,
                    prefix=f"{workbook_path.stem}.",
                    dir=str(workbook_path.parent),
                    delete=False,
                ).name
            )
            try:
                workbook.save(temp_path)
                preserve_workbook_protection(workbook_path, temp_path, log_file)
                os.replace(temp_path, workbook_path)
            finally:
                if temp_path.exists():
                    try:
                        temp_path.unlink()
                    except Exception:
                        pass
            log_event("Workbook save complete.", log_file)
            return
        except Exception as exc:
            if not is_local_conflict_error(exc):
                log_exception("Unhandled workbook save error.", exc, log_file)
                raise
            log_exception("Workbook save conflict detected; will retry.", exc, log_file)
            log_event(
                f"Workbook is still open or locked. Retrying in {retry_delay_minutes} minutes.",
                log_file,
            )
            time.sleep(retry_delay_minutes * 60)
