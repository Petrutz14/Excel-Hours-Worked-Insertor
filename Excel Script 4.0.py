
"""
Pontaj Automation — M&B
Refactored with: #1, #2, #4, #6, #8, #9, #11, #12, #13, #14, #16, #17, #18, #21, #22
"""

import json
import sys
import os
import calendar
import threading
import re
from pathlib import Path  # (#18)
from datetime import datetime
from collections import defaultdict
from typing import Any, Callable, Optional  # (#22)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import win32com.client as win32
import win32process
import win32api
import pythoncom  # (#11) needed for COM in threads
import pywintypes  # (#8) for catching COM errors

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False


# ═══════════════════════════════════════════════════════════════
# CONSTANTS (#21)
# ═══════════════════════════════════════════════════════════════

# Excel layout
SHEET_NAME: str = "CONDICA"
ANGAJATI_SHEET: str = "ANGAJATI"
START_ROW: int = 5
NAME_COLUMN: int = 2  # column B

# Cell positions for date metadata
ANGAJATI_MONTH_CELL: tuple[int, int] = (7, 8)   # H7
CONDICA_YEAR_CELL: tuple[int, int] = (1, 9)     # I1

# Excel constants (#21)
XL_CALCULATION_MANUAL: int = -4135
XL_CALCULATION_AUTOMATIC: int = -4105
XL_FILE_FORMAT_XLS: int = 56  # xlExcel8

# Process termination flag
PROCESS_TERMINATE: int = 1

# Progress phases (#16)
PHASE_COLLECT_PCT: float = 50.0   # 0-50% = collecting
PHASE_WRITE_PCT: float = 50.0     # 50-100% = writing

# Paths (#18)
SCRIPT_DIR: Path = Path(__file__).resolve().parent
SETTINGS_FILE: Path = SCRIPT_DIR / "settings.json"

# Time format validation regex
TIME_RE = re.compile(r"^\d{1,2}[.:]\d{2}$")

DEFAULT_SETTINGS: dict[str, Any] = {
    "last_excel_path": "",
    "last_json_path": "",
    "output_directory": "",
    "log_directory": "",
    "window_geometry": "",  # (#17)
}


# ═══════════════════════════════════════════════════════════════
# SETTINGS
# ═══════════════════════════════════════════════════════════════

def load_settings() -> dict[str, Any]:
    if SETTINGS_FILE.is_file():
        try:
            with SETTINGS_FILE.open("r", encoding="utf-8") as f:
                saved = json.load(f)
            settings = DEFAULT_SETTINGS.copy()
            settings.update(saved)
            return settings
        except Exception:
            return DEFAULT_SETTINGS.copy()
    save_settings(DEFAULT_SETTINGS)
    return DEFAULT_SETTINGS.copy()


def save_settings(settings: dict[str, Any]) -> None:
    try:
        with SETTINGS_FILE.open("w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════════
# NAME NORMALIZATION
# ═══════════════════════════════════════════════════════════════

def normalize_name(s: str) -> str:
    """Locale-safe, whitespace-tolerant key for name lookups."""
    return " ".join(s.strip().casefold().split())


# ═══════════════════════════════════════════════════════════════
# JSON VALIDATION (#6 - duplicate detection added)
# ═══════════════════════════════════════════════════════════════

def validate_json(json_path: Path) -> dict[str, Any]:
    """Validate JSON structure before opening Excel. Fail fast."""
    try:
        with json_path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON syntax: {e}")
    except Exception as e:
        raise ValueError(f"Cannot read JSON file: {e}")

    if not isinstance(data, dict):
        raise ValueError("JSON root must be an object.")
    if "employees" not in data:
        raise ValueError("JSON missing 'employees' key.")
    if not isinstance(data["employees"], list):
        raise ValueError("'employees' must be a list.")
    if len(data["employees"]) == 0:
        raise ValueError("'employees' list is empty.")

    # (#6) Track duplicates
    seen_names: dict[str, int] = {}

    for i, emp in enumerate(data["employees"]):
        label = f"Employee #{i + 1}"
        if not isinstance(emp, dict):
            raise ValueError(f"{label}: must be an object.")
        if "name" not in emp:
            raise ValueError(f"{label}: missing 'name'.")
        if not isinstance(emp["name"], str) or not emp["name"].strip():
            raise ValueError(f"{label}: 'name' must be a non-empty string.")

        # (#6) Duplicate detection
        key = normalize_name(emp["name"])
        if key in seen_names:
            raise ValueError(
                f"Duplicate employee '{emp['name']}' "
                f"(also at #{seen_names[key] + 1})."
            )
        seen_names[key] = i

        if "days" not in emp:
            raise ValueError(f"{label} ({emp['name']}): missing 'days'.")
        if not isinstance(emp["days"], dict):
            raise ValueError(f"{label} ({emp['name']}): 'days' must be an object.")

        for day_str, hours in emp["days"].items():
            try:
                day = int(day_str)
                if day < 1 or day > 31:
                    raise ValueError()
            except (ValueError, TypeError):
                raise ValueError(
                    f"{emp['name']}: invalid day key '{day_str}' (must be 1-31)."
                )
            if not isinstance(hours, str):
                raise ValueError(
                    f"{emp['name']}, day {day_str}: hours must be a string."
                )
            if "-" in hours:
                parts = hours.split("-")
                if len(parts) != 2:
                    raise ValueError(
                        f"{emp['name']}, day {day_str}: "
                        f"invalid hours format '{hours}' (expected 'HH.MM-HH.MM')."
                    )
                entry, exit_time = parts[0].strip(), parts[1].strip()
                if not TIME_RE.match(entry) or not TIME_RE.match(exit_time):
                    raise ValueError(
                        f"{emp['name']}, day {day_str}: "
                        f"invalid time format '{hours}'."
                    )

    return data


# ═══════════════════════════════════════════════════════════════
# LOGGER
# ═══════════════════════════════════════════════════════════════

class Logger:
    def __init__(self, log_dir: Optional[Path] = None) -> None:
        self.lines: list[str] = []
        self.log_dir: Optional[Path] = log_dir
        self.log_path: Optional[Path] = None
        self.enabled: bool = bool(log_dir and log_dir.is_dir())

    def log(self, msg: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.lines.append(f"[{timestamp}] {msg}")

    def log_section(self, title: str) -> None:
        self.lines.append("")
        self.lines.append(f"{'═' * 60}")
        self.lines.append(f"  {title}")
        self.lines.append(f"{'═' * 60}")

    def log_blank(self) -> None:
        self.lines.append("")

    def save(self) -> Optional[Path]:
        if not self.enabled or self.log_dir is None:
            return None
        try:
            timestamp = datetime.now().strftime("%d_%m_%Y-%H-%M-%S")
            self.log_path = self.log_dir / f"MB_Log_{timestamp}.txt"
            with self.log_path.open("w", encoding="utf-8") as f:
                f.write("\n".join(self.lines))
            return self.log_path
        except Exception:
            return None


# ═══════════════════════════════════════════════════════════════
# COLUMN MATH
# ═══════════════════════════════════════════════════════════════

def start_col_for_day(day: int) -> int:
    return 3 + (day - 1) * 7


def entry_col_for_day(day: int) -> int:
    return start_col_for_day(day) + 3


def exit_col_for_day(day: int) -> int:
    return start_col_for_day(day) + 5


# ═══════════════════════════════════════════════════════════════
# CACHED ROW LOOKUP (locale-safe)
# ═══════════════════════════════════════════════════════════════

def build_row_cache(sheet: Any) -> dict[str, int]:
    """Build {normalized_name: row} for O(1) lookups."""
    cache: dict[str, int] = {}
    row = START_ROW
    while True:
        val = sheet.Cells(row, NAME_COLUMN).Value
        if val is None:
            break
        key = normalize_name(str(val))
        if key:
            cache[key] = row
        row += 1
    return cache


# ═══════════════════════════════════════════════════════════════
# EXCEL PROCESS LIFECYCLE
# ═══════════════════════════════════════════════════════════════

def get_excel_pid(excel: Any) -> Optional[int]:
    try:
        hwnd = excel.Hwnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid
    except Exception:
        return None


def force_kill_excel(pid: Optional[int]) -> None:
    if pid is None:
        return
    try:
        handle = win32api.OpenProcess(PROCESS_TERMINATE, False, pid)
        win32api.TerminateProcess(handle, 0)
        win32api.CloseHandle(handle)
    except Exception:
        pass


def safe_quit_excel(excel: Any, pid: Optional[int]) -> None:
    try:
        excel.Quit()
    except Exception:
        pass
    finally:
        force_kill_excel(pid)


# ═══════════════════════════════════════════════════════════════
# OUTPUT PATH HELPER (#18)
# ═══════════════════════════════════════════════════════════════

def get_incremented_path(output_dir: Path, output_name: str) -> Path:
    base = Path(output_name).stem
    ext = Path(output_name).suffix
    counter = 2
    while True:
        new_path = output_dir / f"{base} ({counter}){ext}"
        if not new_path.exists():
            return new_path
        counter += 1


# ═══════════════════════════════════════════════════════════════
# DATE LOGIC
# ═══════════════════════════════════════════════════════════════

def get_expected_month_year() -> tuple[int, int]:
    now = datetime.now()
    year = now.year
    month = now.month - 1
    if month == 0:
        month = 12
        year -= 1
    return month, year


# ═══════════════════════════════════════════════════════════════
# (#9) EXCEL FILE LOCK CHECK
# ═══════════════════════════════════════════════════════════════

def is_excel_locked(path: Path) -> bool:
    """Detect Excel's lock file (~$filename)."""
    lock = path.parent / f"~${path.name}"
    return lock.exists()


# ═══════════════════════════════════════════════════════════════
# WORKING HOURS CHECK
# ═══════════════════════════════════════════════════════════════

def check_employee_hours(condica: Any, row: int, num_days: int) -> bool:
    for day in range(1, num_days + 1):
        entry_val = condica.Cells(row, entry_col_for_day(day)).Value
        exit_val = condica.Cells(row, exit_col_for_day(day)).Value
        if entry_val is not None or exit_val is not None:
            return True
    return False


def get_all_employees(condica: Any) -> list[tuple[str, int]]:
    employees: list[tuple[str, int]] = []
    row = START_ROW
    while True:
        val = condica.Cells(row, NAME_COLUMN).Value
        if val is None:
            break
        employees.append((str(val).strip(), row))
        row += 1
    return employees


# ═══════════════════════════════════════════════════════════════
# (#12) CANCELLATION SUPPORT
# ═══════════════════════════════════════════════════════════════

class CancelledError(Exception):
    """Raised when user cancels mid-run."""
    pass


def check_cancel(cancel_event: Optional[threading.Event]) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Operation cancelled by user.")


# ═══════════════════════════════════════════════════════════════
# CORE AUTOMATION
# ═══════════════════════════════════════════════════════════════

def run_automation(
    excel_path: Path,
    json_path: Path,
    output_dir: Path,
    log_dir: Optional[Path] = None,
    dry_run: bool = False,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    conflict_callback: Optional[Callable[[Path, str, Path], str]] = None,
    date_change_callback: Optional[Callable[[list[str]], None]] = None,
    cancel_event: Optional[threading.Event] = None,  # (#12)
) -> tuple[bool, str, bool, list[str], list[str], list[str]]:
    """
    Returns: (success, result, is_dry, date_changes, with_hours, without_hours)
    """
    logger = Logger(log_dir)
    logger.log_section("PONTAJ AUTOMATION — M&B")
    logger.log(f"Run type: {'DRY RUN' if dry_run else 'LIVE RUN'}")

    logger.log_section("SETTINGS")
    logger.log(f"Excel file:        {excel_path}")
    logger.log(f"JSON file:         {json_path}")
    logger.log(f"Output directory:  {output_dir}")
    logger.log(f"Log directory:     {log_dir if log_dir else 'Not set'}")

    # (#9) Lock file check
    if is_excel_locked(excel_path):
        msg = (
            f"The Excel file appears to be open in another window:\n{excel_path}\n\n"
            f"Please close it before running."
        )
        logger.log(f"LOCK FILE DETECTED: {msg}")
        logger.save()  # (#4)
        return False, msg, False, [], [], []

    # (#4) Pre-validate JSON
    logger.log_section("JSON VALIDATION")
    try:
        data = validate_json(json_path)
        logger.log(f"JSON valid — {len(data['employees'])} employees found.")
    except ValueError as e:
        logger.log(f"JSON VALIDATION FAILED: {e}")
        logger.save()  # (#4)
        return False, str(e), False, [], [], []

    employees = data["employees"]
    total = len(employees)

    # (#2) DispatchEx — fresh isolated process
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    pid = get_excel_pid(excel)
    wb = None
    date_changes: list[str] = []
    with_hours: list[str] = []
    without_hours: list[str] = []

    try:
        # (#8) Wrap COM open with specific error handling
        try:
            wb = excel.Workbooks.Open(str(excel_path.resolve()))
        except pywintypes.com_error as e:
            raise ValueError(
                f"Excel could not open the file. It may be locked, "
                f"corrupted, or already open.\n\nDetails: {e}"
            )

        condica = wb.Worksheets(SHEET_NAME)

        try:
            angajati = wb.Worksheets(ANGAJATI_SHEET)
        except Exception:
            raise ValueError(f"Sheet '{ANGAJATI_SHEET}' not found in the Excel file.")

        # Disable updates / auto-calc
        excel.ScreenUpdating = False
        excel.Calculation = XL_CALCULATION_MANUAL

        check_cancel(cancel_event)

        # ── DATE CHECK ──
        logger.log_section("DATE CHECK")
        expected_month, expected_year = get_expected_month_year()
        logger.log(f"System date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.log(f"Expected month (previous): {expected_month}")
        logger.log(f"Expected year: {expected_year}")

        m_row, m_col = ANGAJATI_MONTH_CELL
        y_row, y_col = CONDICA_YEAR_CELL

        current_month_val = angajati.Cells(m_row, m_col).Value
        current_year_val = condica.Cells(y_row, y_col).Value

        current_month = int(current_month_val) if current_month_val is not None else None
        current_year = int(current_year_val) if current_year_val is not None else None

        logger.log(f"Current ANGAJATI!H7 (month): {current_month}")
        logger.log(f"Current CONDICA!I1  (year):  {current_year}")

        if current_month != expected_month:
            date_changes.append(
                f"Month: ANGAJATI!H7 changed from "
                f"{current_month if current_month is not None else 'EMPTY'} → {expected_month}"
            )
            if not dry_run:
                angajati.Cells(m_row, m_col).Value = expected_month
            logger.log(f"CHANGED month: {current_month} → {expected_month}")
        else:
            logger.log("Month OK — no change needed.")

        if current_year != expected_year:
            date_changes.append(
                f"Year: CONDICA!I1 changed from "
                f"{current_year if current_year is not None else 'EMPTY'} → {expected_year}"
            )
            if not dry_run:
                condica.Cells(y_row, y_col).Value = expected_year
            logger.log(f"CHANGED year: {current_year} → {expected_year}")
        else:
            logger.log("Year OK — no change needed.")

        if date_changes and date_change_callback:
            date_change_callback(date_changes)

        month_str = str(expected_month).zfill(2)
        year_str = str(expected_year)

        num_days = calendar.monthrange(expected_year, expected_month)[1]
        logger.log(f"Days in month: {num_days}")

        # ── ROW CACHE ──
        logger.log_section("ROW CACHE")
        row_cache = build_row_cache(condica)
        logger.log(f"Cached {len(row_cache)} employee rows.")

        check_cancel(cancel_event)

        # ── COLLECT WRITES (Phase 1: 0-50%) (#16) ──
        logger.log_section("EMPLOYEE PROCESSING")
        entries_count = 0
        writes: list[tuple[int, int, str]] = []

        for i, emp in enumerate(employees):
            check_cancel(cancel_event)

            key = normalize_name(emp["name"])
            row = row_cache.get(key)
            if not row:
                raise ValueError(f"Employee name not found: '{emp['name']}'")

            emp_entries = 0
            for day_str, hours in emp["days"].items():
                if "-" not in hours:
                    continue
                day = int(day_str)
                entry, exit_time = (s.strip() for s in hours.split("-"))
                writes.append((row, entry_col_for_day(day), entry))
                writes.append((row, exit_col_for_day(day), exit_time))
                entries_count += 1
                emp_entries += 1

            logger.log(f"  {emp['name']} — {emp_entries} entries (row {row})")

            # (#16) Phase 1 progress
            if progress_callback:
                pct = ((i + 1) / total) * PHASE_COLLECT_PCT
                progress_callback(pct, f"Collecting {i + 1}/{total}: {emp['name']}")

        logger.log_blank()
        logger.log(f"Total entries to write: {entries_count}")
        logger.log(f"Total cell writes: {len(writes)}")

        check_cancel(cancel_event)

        # ── BATCH WRITE (Phase 2: 50-100%) (#1, #16) ──
        if not dry_run and writes:
            logger.log("Writing cells individually...")

            # Group by row
            row_writes: dict[int, dict[int, str]] = defaultdict(dict)
            for r, c, v in writes:
                row_writes[r][c] = v

            row_items = list(row_writes.items())
            row_count = len(row_items)

            for idx, (r, col_vals) in enumerate(row_items):
                check_cancel(cancel_event)

                if not col_vals:
                    continue

                for c, v in col_vals.items():
                    condica.Cells(r, c).Value = v

                # (#16) Phase 2 progress
                if progress_callback:
                    pct = PHASE_COLLECT_PCT + ((idx + 1) / row_count) * PHASE_WRITE_PCT
                    progress_callback(pct, f"Writing row {idx + 1}/{row_count}")

            logger.log("Batch write complete.")

        check_cancel(cancel_event)

        # ── WORKING HOURS CHECK ──
        logger.log_section("WORKING HOURS CHECK")
        all_employees = get_all_employees(condica)
        for name, row in all_employees:
            if check_employee_hours(condica, row, num_days):
                with_hours.append(name)
            else:
                without_hours.append(name)

        logger.log(f"Employees WITH working hours ({len(with_hours)}):")
        for name in with_hours:
            logger.log(f"  ✓ {name}")
        logger.log_blank()
        logger.log(f"Employees WITHOUT working hours ({len(without_hours)}):")
        for name in without_hours:
            logger.log(f"  ✗ {name}")

        excel.ScreenUpdating = True
        excel.Calculation = XL_CALCULATION_AUTOMATIC
        excel.Calculate()

        # ── DRY RUN PATH ──
        if dry_run:
            logger.log_section("DRY RUN SUMMARY")
            logger.log("No files were modified or saved.")

            wb.Close(SaveChanges=False)
            wb = None
            safe_quit_excel(excel, pid)

            summary = (
                f"Dry Run Complete\n\n"
                f"Employees found: {total}\n"
                f"Entries to write: {entries_count}\n"
                f"Days in month: {num_days}\n"
                f"Output would be:\n"
                f"CONDICA_PONTAJ_{month_str}.{year_str} - M&B.xls"
            )
            if date_changes:
                summary += "\n\nDate changes that WOULD be made:\n"
                summary += "\n".join(f"  • {c}" for c in date_changes)
            summary += f"\n\nEmployees with hours: {len(with_hours)}"
            summary += f"\nEmployees without hours: {len(without_hours)}"

            logger.save()  # (#4)
            return True, summary, True, date_changes, with_hours, without_hours

        # ── LIVE SAVE ──
        logger.log_section("OUTPUT")
        output_name = f"CONDICA_PONTAJ_{month_str}.{year_str} - M&B.xls"
        output_path = output_dir / output_name

        if output_path.exists() and conflict_callback:
            decision = conflict_callback(output_path, output_name, output_dir)
            if decision == "cancel":
                logger.log("User cancelled due to file conflict.")
                wb.Close(SaveChanges=False)
                wb = None
                safe_quit_excel(excel, pid)
                logger.save()
                return False, "Operation cancelled by user.", False, date_changes, with_hours, without_hours
            elif decision == "increment":
                output_path = get_incremented_path(output_dir, output_name)
                logger.log(f"Auto-incremented to: {output_path.name}")

        logger.log(f"Saving to: {output_path}")
        try:
            wb.SaveAs(str(output_path.resolve()), FileFormat=XL_FILE_FORMAT_XLS)
        except pywintypes.com_error as e:  # (#8)
            raise ValueError(f"Failed to save Excel file: {e}")

        wb.Close()
        wb = None
        safe_quit_excel(excel, pid)

        logger.log("File saved successfully.")
        logger.log_section("COMPLETE")
        logger.log("Automation finished successfully.")
        logger.save()

        return True, str(output_path), False, date_changes, with_hours, without_hours

    except CancelledError as e:  # (#12)
        logger.log_section("CANCELLED")
        logger.log(str(e))
        logger.save()
        try:
            excel.ScreenUpdating = True
            excel.Calculation = XL_CALCULATION_AUTOMATIC
        except Exception:
            pass
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        safe_quit_excel(excel, pid)
        return False, "Operation cancelled by user.", False, date_changes, with_hours, without_hours

    except Exception as e:
        logger.log_section("ERROR")
        logger.log(f"FATAL: {e}")
        logger.log("Operation aborted — nothing was saved.")
        logger.save()

        try:
            excel.ScreenUpdating = True
            excel.Calculation = XL_CALCULATION_AUTOMATIC
        except Exception:
            pass
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        safe_quit_excel(excel, pid)
        return False, str(e), False, date_changes, with_hours, without_hours


# ═══════════════════════════════════════════════════════════════
# OS HELPERS
# ═══════════════════════════════════════════════════════════════

def open_folder(path: Path) -> None:
    try:
        os.startfile(str(path))
    except Exception:
        pass


def open_file(path: Path) -> None:
    try:
        os.startfile(str(path))
    except Exception:
        pass


def clean_dropped_path(raw: str) -> str:
    path = raw.strip()
    if path.startswith("{") and path.endswith("}"):
        path = path[1:-1]
    if path.startswith('"') and path.endswith('"'):
        path = path[1:-1]
    return path


# ═══════════════════════════════════════════════════════════════
# GUI
# ═══════════════════════════════════════════════════════════════

class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Pontaj Automation — M&B")

        self.settings = load_settings()

        # (#17) Restore window geometry
        geom = self.settings.get("window_geometry", "")
        if geom:
            try:
                self.root.geometry(geom)
            except Exception:
                self.root.geometry("620x680")
        else:
            self.root.geometry("620x680")
        self.root.resizable(False, False)

        # (#17) Save geometry on close
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # State
        self.excel_path = tk.StringVar(value=self.settings.get("last_excel_path", ""))
        self.json_path = tk.StringVar(value=self.settings.get("last_json_path", ""))
        self.output_dir = tk.StringVar(value=self.settings.get("output_directory", ""))
        self.log_dir = tk.StringVar(value=self.settings.get("log_directory", ""))
        self.dry_run = tk.BooleanVar(value=False)

        self.last_output_path: Optional[Path] = None
        self.last_log_path: Optional[Path] = None

        # (#11, #12) Threading state
        self.worker_thread: Optional[threading.Thread] = None
        self.cancel_event: Optional[threading.Event] = None

        self._build_ui()

        if HAS_DND:
            self.setup_drag_and_drop()
        else:
            tk.Label(
                root,
                text="💡 Install 'tkinterdnd2' (pip install tkinterdnd2) to enable drag & drop",
                font=("Segoe UI", 8), fg="gray"
            ).pack(side="bottom", pady=(0, 5))

    # ── UI BUILD ──
    def _build_ui(self) -> None:
        tk.Label(
            self.root, text="Pontaj Automation", font=("Segoe UI", 16, "bold")
        ).pack(pady=(15, 10))

        # Excel
        f1 = tk.Frame(self.root); f1.pack(fill="x", padx=20, pady=4)
        tk.Label(f1, text="Excel File (.xls):", width=16, anchor="w").pack(side="left")
        self.excel_entry = tk.Entry(f1, textvariable=self.excel_path, width=38)
        self.excel_entry.pack(side="left", padx=5)
        tk.Button(f1, text="Browse", command=self.browse_excel).pack(side="left")

        # JSON
        f2 = tk.Frame(self.root); f2.pack(fill="x", padx=20, pady=4)
        tk.Label(f2, text="JSON File:", width=16, anchor="w").pack(side="left")
        self.json_entry = tk.Entry(f2, textvariable=self.json_path, width=38)
        self.json_entry.pack(side="left", padx=5)
        tk.Button(f2, text="Browse", command=self.browse_json).pack(side="left")

        # Output
        f3 = tk.Frame(self.root); f3.pack(fill="x", padx=20, pady=4)
        tk.Label(f3, text="Output Directory:", width=16, anchor="w").pack(side="left")
        tk.Entry(f3, textvariable=self.output_dir, width=38).pack(side="left", padx=5)
        tk.Button(f3, text="Browse", command=self.browse_output).pack(side="left")

        ttk.Separator(self.root, orient="horizontal").pack(fill="x", padx=20, pady=8)

        # Log dir
        f4 = tk.Frame(self.root); f4.pack(fill="x", padx=20, pady=4)
        tk.Label(f4, text="Log Directory:", width=16, anchor="w",
                 font=("Segoe UI", 9), fg="gray").pack(side="left")
        tk.Entry(f4, textvariable=self.log_dir, width=38).pack(side="left", padx=5)
        tk.Button(f4, text="Browse", command=self.browse_log).pack(side="left")
        tk.Label(f4, text="(optional)", font=("Segoe UI", 8), fg="gray").pack(side="left", padx=3)

        ttk.Separator(self.root, orient="horizontal").pack(fill="x", padx=20, pady=8)

        tk.Checkbutton(
            self.root, text="Dry Run (preview only — no changes saved)",
            variable=self.dry_run, font=("Segoe UI", 10)
        ).pack(pady=(0, 5))

        # Progress
        pf = tk.Frame(self.root); pf.pack(fill="x", padx=20, pady=5)
        self.progress_bar = ttk.Progressbar(
            pf, orient="horizontal", length=560, mode="determinate", maximum=100
        )
        self.progress_bar.pack()
        self.progress_label = tk.Label(pf, text="", font=("Segoe UI", 9), fg="gray")
        self.progress_label.pack(pady=(2, 0))

        # (#16) Phase indicator
        self.phase_label = tk.Label(pf, text="", font=("Segoe UI", 8), fg="gray")
        self.phase_label.pack()

        # Buttons
        bf = tk.Frame(self.root); bf.pack(pady=10)

        self.run_btn = tk.Button(
            bf, text="▶  Run", font=("Segoe UI", 11, "bold"),
            bg="#4CAF50", fg="white", activebackground="#45a049",
            command=self.run, width=12, height=2
        )
        self.run_btn.pack(side="left", padx=5)

        # (#12) Cancel button
        self.cancel_btn = tk.Button(
            bf, text="✖ Cancel", font=("Segoe UI", 10),
            bg="#f44336", fg="white", activebackground="#d32f2f",
            command=self.cancel, width=10, height=2,
            state="disabled"
        )
        self.cancel_btn.pack(side="left", padx=5)

        self.open_folder_btn = tk.Button(
            bf, text="📂 Output", font=("Segoe UI", 10),
            command=self.open_output_folder, width=10, height=2,
            state="disabled"
        )
        self.open_folder_btn.pack(side="left", padx=5)

        self.open_log_btn = tk.Button(
            bf, text="📄 Log", font=("Segoe UI", 10),
            command=self.open_log_file, width=8, height=2,
            state="disabled"
        )
        self.open_log_btn.pack(side="left", padx=5)

        self.status = tk.Label(
            self.root, text="Select files and click Run.",
            font=("Segoe UI", 10), wraplength=560, justify="center"
        )
        self.status.pack(pady=(0, 10))

    # ── DRAG & DROP ──
    def setup_drag_and_drop(self) -> None:
        def drop_excel(event: Any) -> None:
            path = clean_dropped_path(event.data)
            if path.lower().endswith((".xls", ".xlsx")):
                self.excel_path.set(path)
                if not self.output_dir.get().strip():
                    self.output_dir.set(str(Path(path).parent))

        def drop_json(event: Any) -> None:
            path = clean_dropped_path(event.data)
            if path.lower().endswith(".json"):
                self.json_path.set(path)

        self.excel_entry.drop_target_register(DND_FILES)
        self.excel_entry.dnd_bind("<<Drop>>", drop_excel)
        self.json_entry.drop_target_register(DND_FILES)
        self.json_entry.dnd_bind("<<Drop>>", drop_json)

        tk.Label(
            self.root,
            text="💡 You can drag & drop .xls and .json files onto the fields above",
            font=("Segoe UI", 8), fg="gray"
        ).pack(side="bottom", pady=(0, 5))

    # ── BROWSE ──
    def browse_excel(self) -> None:
        cur = self.excel_path.get()
        initial = str(Path(cur).parent) if cur else ""
        path = filedialog.askopenfilename(
            title="Select Excel File", initialdir=initial,
            filetypes=[("Excel Files", "*.xls *.xlsx"), ("All Files", "*.*")]
        )
        if path:
            self.excel_path.set(path)
            if not self.output_dir.get().strip():
                self.output_dir.set(str(Path(path).parent))

    def browse_json(self) -> None:
        cur = self.json_path.get()
        initial = str(Path(cur).parent) if cur else ""
        path = filedialog.askopenfilename(
            title="Select JSON File", initialdir=initial,
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if path:
            self.json_path.set(path)

    def browse_output(self) -> None:
        path = filedialog.askdirectory(
            title="Select Output Directory", initialdir=self.output_dir.get() or ""
        )
        if path:
            self.output_dir.set(path)

    def browse_log(self) -> None:
        path = filedialog.askdirectory(
            title="Select Log Directory", initialdir=self.log_dir.get() or ""
        )
        if path:
            self.log_dir.set(path)

    # ── BUTTONS ──
    def open_output_folder(self) -> None:
        if self.last_output_path:
            open_folder(self.last_output_path.parent)

    def open_log_file(self) -> None:
        if self.last_log_path and self.last_log_path.is_file():
            open_file(self.last_log_path)

    # ── CONFLICT DIALOG (must run on main thread) ──
    def handle_conflict(self, output_path: Path, output_name: str, output_dir: Path) -> str:
        result_holder: dict[str, str] = {}
        done = threading.Event()

        def ask() -> None:
            res = messagebox.askyesnocancel(
                "File Already Exists",
                f"'{output_name}' already exists in:\n{output_dir}\n\n"
                f"Yes = Overwrite\nNo = Save as new copy (auto-increment)\nCancel = Abort"
            )
            if res is True:
                result_holder["v"] = "overwrite"
            elif res is False:
                result_holder["v"] = "increment"
            else:
                result_holder["v"] = "cancel"
            done.set()

        self.root.after(0, ask)
        done.wait()
        return result_holder["v"]

    def handle_date_change(self, changes: list[str]) -> None:
        # Marshal to main thread
        msg = "The following date changes were made:\n\n"
        for c in changes:
            msg += f"  • {c}\n"
        self.root.after(0, lambda: messagebox.showinfo("Date Updated", msg))

    # ── PROGRESS (#13, #16) ──
    def update_progress(self, pct: float, label: str) -> None:
        # Called from worker thread — marshal to main
        def apply() -> None:
            self.progress_bar["value"] = pct
            self.progress_label.config(text=label)
            if pct < PHASE_COLLECT_PCT:
                self.phase_label.config(text="Phase 1/2: Collecting writes")
            else:
                self.phase_label.config(text="Phase 2/2: Writing to Excel")
            # (#13) safer than update()
            self.root.update_idletasks()

        self.root.after(0, apply)

    # ── SETTINGS ──
    def save_current_settings(self) -> None:
        self.settings["last_excel_path"] = self.excel_path.get().strip()
        self.settings["last_json_path"] = self.json_path.get().strip()
        self.settings["output_directory"] = self.output_dir.get().strip()
        self.settings["log_directory"] = self.log_dir.get().strip()
        self.settings["window_geometry"] = self.root.geometry()  # (#17)
        save_settings(self.settings)

    def _on_close(self) -> None:
        # (#17) Save geometry before exit
        if self.worker_thread and self.worker_thread.is_alive():
            if not messagebox.askyesno(
                "Quit",
                "An automation is still running. Cancel it and quit?"
            ):
                return
            if self.cancel_event:
                self.cancel_event.set()
        self.save_current_settings()
        self.root.destroy()

    def reset_progress(self) -> None:
        self.progress_bar["value"] = 0
        self.progress_label.config(text="")
        self.phase_label.config(text="")
        self.open_folder_btn.config(state="disabled")
        self.open_log_btn.config(state="disabled")
        self.last_output_path = None
        self.last_log_path = None

    # ── CANCEL (#12) ──
    def cancel(self) -> None:
        if self.cancel_event:
            self.cancel_event.set()
            self.status.config(text="⏹ Cancelling...", fg="orange")
            self.cancel_btn.config(state="disabled")

    # ── RUN (#11) ──
    def run(self) -> None:
        excel_file = self.excel_path.get().strip()
        json_file = self.json_path.get().strip()
        output_dir = self.output_dir.get().strip()
        log_dir = self.log_dir.get().strip() or None
        is_dry = self.dry_run.get()

        # Validate
        if not excel_file:
            messagebox.showwarning("Missing File", "Please select an Excel file.")
            return
        if not json_file:
            messagebox.showwarning("Missing File", "Please select a JSON file.")
            return
        if not output_dir and not is_dry:
            messagebox.showwarning("Missing Directory", "Please select an output directory.")
            return

        excel_path = Path(excel_file)
        json_path = Path(json_file)
        output_path = Path(output_dir) if output_dir else Path()
        log_path = Path(log_dir) if log_dir else None

        if not excel_path.is_file():
            messagebox.showerror("Error", f"Excel file not found:\n{excel_path}")
            return
        if not json_path.is_file():
            messagebox.showerror("Error", f"JSON file not found:\n{json_path}")
            return
        if not is_dry and not output_path.is_dir():
            messagebox.showerror("Error", f"Output directory not found:\n{output_path}")
            return
        if log_path and not log_path.is_dir():
            messagebox.showerror("Error", f"Log directory not found:\n{log_path}")
            return

        self.save_current_settings()
        self.reset_progress()

        # (#11) Set up cancel + thread
        self.cancel_event = threading.Event()
        self.run_btn.config(
            state="disabled",
            text="🔍 Dry Run..." if is_dry else "⏳ Processing..."
        )
        self.cancel_btn.config(state="normal")
        self.status.config(text="Working... please wait.", fg="orange")

        self.worker_thread = threading.Thread(
            target=self._worker,
            args=(excel_path, json_path, output_path, log_path, is_dry),
            daemon=True,
        )
        self.worker_thread.start()

    # ── (#11) WORKER THREAD ──
    def _worker(
        self,
        excel_path: Path,
        json_path: Path,
        output_path: Path,
        log_path: Optional[Path],
        is_dry: bool,
    ) -> None:
        # (#11) MUST initialize COM in this thread
        pythoncom.CoInitialize()
        try:
            success, result, is_dry_result, date_changes, with_hours, without_hours = run_automation(
                excel_path, json_path, output_path,
                log_dir=log_path, dry_run=is_dry,
                progress_callback=self.update_progress,
                conflict_callback=self.handle_conflict,
                date_change_callback=self.handle_date_change,
                cancel_event=self.cancel_event,
            )
        finally:
            pythoncom.CoUninitialize()

        # Marshal results back to main thread
        self.root.after(
            0,
            lambda: self._on_complete(
                success, result, is_dry_result, date_changes, with_hours, without_hours, log_path
            ),
        )

    # ── COMPLETION (main thread) ──
    def _on_complete(
        self,
        success: bool,
        result: str,
        is_dry_result: bool,
        date_changes: list[str],
        with_hours: list[str],
        without_hours: list[str],
        log_path: Optional[Path],
    ) -> None:
        # Reset run/cancel buttons
        self.run_btn.config(state="normal", text="▶  Run")
        self.cancel_btn.config(state="disabled")
        self.cancel_event = None
        self.worker_thread = None

        # Find latest log
        if log_path and log_path.is_dir():
            logs = sorted(
                [f for f in os.listdir(log_path) if f.startswith("MB_Log_")],
                reverse=True,
            )
            if logs:
                self.last_log_path = log_path / logs[0]
                self.open_log_btn.config(state="normal")

        # ── Show results ──
        if success:
            if is_dry_result:
                self.progress_bar["value"] = 100
                self.progress_label.config(text="Dry run — no files were modified.")
                self.phase_label.config(text="")
                self.status.config(text="🔍 Dry Run completed.", fg="blue")

                # (#14) Surface empty-hours warning in dialog
                msg = result
                if without_hours:
                    msg += self._format_empty_hours_warning(without_hours)

                messagebox.showinfo("Dry Run Result", msg)
            else:
                self.last_output_path = Path(result)
                self.open_folder_btn.config(state="normal")

                basename = self.last_output_path.name
                self.progress_label.config(text=f"Saved: {basename}")
                self.phase_label.config(text="")

                status_text = f"✅ Saved as:\n{basename}"
                if date_changes:
                    status_text += "\n📅 Date was updated."
                if without_hours:
                    status_text += f"\n⚠️ {len(without_hours)} employee(s) without hours."
                self.status.config(text=status_text, fg="green")

                msg = f"File saved as:\n{result}"
                if date_changes:
                    msg += "\n\nDate changes made:\n"
                    msg += "\n".join(f"  • {c}" for c in date_changes)

                # (#14) Surface empty-hours warning in dialog
                if without_hours:
                    msg += self._format_empty_hours_warning(without_hours)

                messagebox.showinfo("Success", msg)
        else:
            self.progress_label.config(text="Failed.")
            self.phase_label.config(text="")
            self.status.config(text=f"❌ Error: {result}", fg="red")
            messagebox.showerror(
                "Error — Aborted",
                f"Nothing was saved.\n\n{result}",
            )

    # ── (#14) Empty-hours formatter ──
    @staticmethod
    def _format_empty_hours_warning(without_hours: list[str]) -> str:
        msg = f"\n\n⚠️  {len(without_hours)} employee(s) have NO hours:\n"
        preview = without_hours[:10]
        msg += "\n".join(f"  • {n}" for n in preview)
        if len(without_hours) > 10:
            msg += f"\n  ... and {len(without_hours) - 10} more (see log)"
        return msg


# ═══════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = App(root)
    root.mainloop()
