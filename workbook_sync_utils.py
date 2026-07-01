from __future__ import annotations

import ctypes
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

WINDOWS_LOCK_ERRORS = {5, 32, 33}


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
            workbook.save(workbook_path)
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
