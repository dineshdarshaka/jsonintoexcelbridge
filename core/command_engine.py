"""
command_engine.py
-----------------
VBA-style cell-level command parser for the Local Bridge Service.

Cell addressing uses standard Excel notation:
  A1  = column A, row 1
  B5  = column B, row 5

Supported commands (case-insensitive, optional / prefix):

  --- Cell Read/Write ---
  GET   A1              - Read a single cell
  SET   A1 = value      - Write value to a cell
  RANGE A1:C10          - Read a rectangular range

  --- Navigation ---
  LASTROW  A            - Last used row number in column A
  LASTCOL  1            - Last used column letter in row 1

  --- Row Operations ---
  INSERTROW  5          - Insert blank row at row 5 (shift down)
  DELETEROW  5          - Delete row 5 (shift up)
  CLEARROW   5          - Clear contents of row 5 (keep row)

  --- Range Clear ---
  CLEAR   A1:C10        - Clear contents of a range
  CLEARCOL  A           - Clear entire column A
  CLR                    - Clear ALL rows (keep columns)

  --- Column Operations ---
  COLS                   - List all column names
  ADDCOL   NAME          - Add a new empty column
  RMVCOL   NAME          - Delete a column and all its data
  RENCOL   OLD  NEW      - Rename a column

  --- Display ---
  SHOW                   - Show entire sheet
  HELP                   - Show this help
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pandas as pd


# ---------------------------------------------------------------------------
# Cell address helpers
# ---------------------------------------------------------------------------

def _parse_cell(addr: str) -> tuple[int, int]:
    """Parse Excel cell like 'A1' or 'B5'. Returns (row, col) both 0-based."""
    addr = addr.strip().upper()
    m = re.match(r"^([A-Z]+)(\d+)$", addr)
    if not m:
        raise ValueError(f"Invalid cell: '{addr}'. Use A1 style (e.g. B5).")
    col_letters = m.group(1)
    row = int(m.group(2)) - 1
    col = 0
    for ch in col_letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    col -= 1
    return row, col


def _col_index_to_letter(col: int) -> str:
    """Convert 0-based column index to Excel letters."""
    result = ""
    while col >= 0:
        result = chr(col % 26 + ord("A")) + result
        col = col // 26 - 1
    return result


def _col_letter_to_index(letter: str) -> int:
    """Convert column letter(s) like 'A', 'AA' to 0-based index."""
    letter = letter.strip().upper()
    col = 0
    for ch in letter:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return col - 1


def _parse_range(range_str: str) -> tuple[int, int, int, int]:
    """Parse 'A1:C10' → (row_start, col_start, row_end, col_end) all 0-based."""
    parts = range_str.strip().upper().split(":")
    if len(parts) != 2:
        raise ValueError(f"Invalid range: '{range_str}'. Use A1:C10 format.")
    r1, c1 = _parse_cell(parts[0])
    r2, c2 = _parse_cell(parts[1])
    return (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def _re_extract_value(original: str, keyword: str) -> str:
    """Extract the value portion from original text after 'keyword addr = '."""
    # Strip leading / if present
    txt = original.strip()
    if txt.startswith("/"):
        txt = txt[1:].strip()
    # Find the = after the keyword and address
    idx = txt.index("=")
    return txt[idx + 1:].strip()


# ---------------------------------------------------------------------------
# Regex patterns (order matters — more specific first)
# ---------------------------------------------------------------------------

_RE_SET = re.compile(r"^SET\s+([A-Z]+\d+)\s*=\s*(.+)$", re.IGNORECASE)
_RE_GET = re.compile(r"^GET\s+([A-Z]+\d+)\s*$", re.IGNORECASE)
_RE_RANGE = re.compile(r"^RANGE\s+([A-Z]+\d+:[A-Z]+\d+)\s*$", re.IGNORECASE)
_RE_LASTROW = re.compile(r"^LASTROW\s+([A-Z]+)\s*$", re.IGNORECASE)
_RE_LASTCOL = re.compile(r"^LASTCOL\s+(\d+)\s*$", re.IGNORECASE)
_RE_INSERTROW = re.compile(r"^INSERTROW\s+(\d+)\s*$", re.IGNORECASE)
_RE_DELETEROW = re.compile(r"^DELETEROW\s+(\d+)\s*$", re.IGNORECASE)
_RE_CLEARROW = re.compile(r"^CLEARROW\s+(\d+)\s*$", re.IGNORECASE)
_RE_CLEAR = re.compile(r"^CLEAR\s+([A-Z]+\d+:[A-Z]+\d+)\s*$", re.IGNORECASE)
_RE_CLEARCOL = re.compile(r"^CLEARCOL\s+([A-Z]+)\s*$", re.IGNORECASE)
_RE_ADDCOL = re.compile(r"^ADDCOL\s+(\w+)\s*$", re.IGNORECASE)
_RE_RMVCOL = re.compile(r"^RMVCOL\s+(\w+)\s*$", re.IGNORECASE)
_RE_RENCOL = re.compile(r"^RENCOL\s+(\w+)\s+(\w+)\s*$", re.IGNORECASE)
_RE_SHOW = re.compile(r"^SHOW\s*$", re.IGNORECASE)
_RE_COLS = re.compile(r"^COLS\s*$", re.IGNORECASE)
_RE_CLR = re.compile(r"^CLR\s*$", re.IGNORECASE)
_RE_HELP = re.compile(r"^HELP\s*$", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Result
# ---------------------------------------------------------------------------

class CommandResult:
    """Result of executing a command."""

    def __init__(
        self,
        success: bool,
        action: str,
        message: str,
        rows_affected: int = 0,
        data: list[dict[str, Any]] | None = None,
        value: Any = None,
    ) -> None:
        self.success = success
        self.action = action
        self.message = message
        self.rows_affected = rows_affected
        self.data = data or []
        self.value = value

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "success": self.success,
            "action": self.action,
            "message": self.message,
            "rows_affected": self.rows_affected,
        }
        if self.value is not None:
            result["value"] = self.value
        if self.data:
            result["data"] = self.data
            result["row_count"] = len(self.data)
        return result


# ---------------------------------------------------------------------------
# Engine
# ---------------------------------------------------------------------------

class CommandEngine:
    """VBA-style cell command engine for Excel manipulation."""

    HELP_TEXT = r"""
📋 VBA Cell Commands

── 🔍 Read ──
  GET A1               Read cell A1
  RANGE A1:C10         Read rectangular range

── ✏️ Write ──
  SET A1 = value       Write value to cell (numbers auto-detected)

── 🧭 Navigate ──
  LASTROW A            Last used row in column A
  LASTCOL 1            Last used column in row 1

── 📝 Rows ──
  INSERTROW 5          Insert blank row at 5 (shift down)
  DELETEROW 5          Delete row 5 (shift up)
  CLEARROW 5           Clear row 5 contents

── 🧹 Clear ──
  CLEAR A1:C10         Clear range contents
  CLEARCOL A           Clear entire column A
  CLR                  Clear ALL rows (keep columns)

── 🏷️ Columns ──
  COLS                 List all columns (A=data, B=name, ...)
  ADDCOL Status        Add new column (auto-assigned letter)
  RMVCOL Temp          Delete a column
  RENCOL Old New       Rename a column

── 📊 ──
  SHOW                 Show entire sheet
  HELP                 Show this help

Prefix with / is optional. All commands case-insensitive.
"""

    def __init__(self, excel_path: Path, sheet_name: str = "Sheet1") -> None:
        self._excel_path = excel_path
        self._sheet_name = sheet_name

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    # Read-only commands that don't modify the Excel file
    _READ_ONLY_COMMANDS = {"HELP", "SHOW", "GET", "RANGE", "LASTROW", "LASTCOL", "COLS"}

    @staticmethod
    def is_read_only(command_text: str) -> bool:
        """Check if a command is read-only (doesn't modify the Excel file)."""
        text = command_text.strip()
        if text.startswith("/"):
            text = text[1:].strip()
        if not text:
            return True  # empty command is harmless
        # Extract the command keyword (first word)
        keyword = text.split()[0].upper() if text.split() else ""
        return keyword in CommandEngine._READ_ONLY_COMMANDS

    def execute(self, command_text: str) -> CommandResult:
        """Parse and execute a single command string."""
        text = command_text.strip()
        if not text:
            return CommandResult(False, "NONE", "Empty command.")
        if text.startswith("/"):
            text = text[1:].strip()
        try:
            return self._dispatch(text.upper(), text)
        except Exception as exc:
            return CommandResult(False, "ERROR", f"Failed: {exc!s}")

    # ------------------------------------------------------------------
    # Load / Save
    # ------------------------------------------------------------------

    def _load(self) -> pd.DataFrame:
        if not self._excel_path.is_file():
            return pd.DataFrame()
        try:
            df = pd.read_excel(self._excel_path, sheet_name=self._sheet_name)
            df = df.where(pd.notna(df), None)
            # Convert all columns to object dtype to allow mixed types
            for col in df.columns:
                df[col] = df[col].astype(object)
            return df
        except Exception:
            return pd.DataFrame()

    def _save(self, df: pd.DataFrame) -> None:
        import os as _os
        self._excel_path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self._excel_path.with_suffix(".tmp.xlsx")
        df.to_excel(tmp, sheet_name=self._sheet_name, index=False)
        _os.replace(tmp, self._excel_path)

    def _to_records(self, df: pd.DataFrame) -> list[dict[str, Any]]:
        return json.loads(df.fillna("").to_json(orient="records"))

    def _ensure_cols(self, df: pd.DataFrame, col: int) -> pd.DataFrame:
        while len(df.columns) <= col:
            df[f"_Col{len(df.columns)}"] = None
        return df

    def _ensure_rows(self, df: pd.DataFrame, row: int) -> pd.DataFrame:
        while len(df) <= row:
            df = pd.concat([df, pd.DataFrame([{c: None for c in df.columns}])], ignore_index=True)
        return df

    # ------------------------------------------------------------------
    # Dispatch
    # ------------------------------------------------------------------

    def _dispatch(self, upper: str, original: str) -> CommandResult:
        m = _RE_HELP.match(upper)
        if m: return self._handle_help()
        m = _RE_COLS.match(upper)
        if m: return self._handle_cols()
        m = _RE_CLR.match(upper)
        if m: return self._handle_clr()
        m = _RE_SHOW.match(upper)
        if m: return self._handle_show()
        m = _RE_GET.match(upper)
        if m: return self._handle_get(m.group(1))
        m = _RE_SET.match(upper)
        if m: return self._handle_set(m.group(1), _re_extract_value(original, "SET"))
        m = _RE_RANGE.match(upper)
        if m: return self._handle_range(m.group(1))
        m = _RE_LASTROW.match(upper)
        if m: return self._handle_lastrow(m.group(1))
        m = _RE_LASTCOL.match(upper)
        if m: return self._handle_lastcol(int(m.group(1)))
        m = _RE_INSERTROW.match(upper)
        if m: return self._handle_insertrow(int(m.group(1)))
        m = _RE_DELETEROW.match(upper)
        if m: return self._handle_deleterow(int(m.group(1)))
        m = _RE_CLEARROW.match(upper)
        if m: return self._handle_clearrow(int(m.group(1)))
        m = _RE_CLEAR.match(upper)
        if m: return self._handle_clear(m.group(1))
        m = _RE_CLEARCOL.match(upper)
        if m: return self._handle_clearcol(m.group(1))
        m = _RE_ADDCOL.match(upper)
        if m: return self._handle_addcol(m.group(1))
        m = _RE_RMVCOL.match(upper)
        if m: return self._handle_rmvcol(m.group(1))
        m = _RE_RENCOL.match(upper)
        if m: return self._handle_rencol(m.group(1), m.group(2))
        return CommandResult(False, "UNKNOWN", f"Unknown command: {original[:80]}\nType HELP.")

    # ------------------------------------------------------------------
    # Handlers
    # ------------------------------------------------------------------

    def _handle_help(self):
        return CommandResult(True, "HELP", self.HELP_TEXT)

    def _handle_cols(self):
        df = self._load()
        mapping = {_col_index_to_letter(i): name for i, name in enumerate(df.columns)}
        return CommandResult(True, "COLS", f"{len(mapping)} column(s).", data=[mapping])

    def _handle_show(self):
        df = self._load()
        rows = self._to_records(df)
        return CommandResult(True, "SHOW", f"{len(rows)} row(s).", data=rows, rows_affected=len(rows))

    def _handle_get(self, addr: str):
        row, col = _parse_cell(addr)
        df = self._load()
        df = self._ensure_cols(df, col)
        if row >= len(df):
            return CommandResult(True, "GET", f"{addr} = empty", value=None)
        val = df.iloc[row, col]
        if pd.isna(val):
            val = None
        elif hasattr(val, "item"):
            val = val.item()
        return CommandResult(True, "GET", f"{addr} = {val!r}", value=val)

    def _handle_set(self, addr: str, raw: str):
        row, col = _parse_cell(addr)
        df = self._load()
        df = self._ensure_cols(df, col)
        df = self._ensure_rows(df, row)
        val: Any = raw
        try:
            val = float(raw) if "." in raw else int(raw)
        except ValueError:
            pass
        df.iloc[row, col] = val
        self._save(df)
        letter = _col_index_to_letter(col)
        return CommandResult(True, "SET", f"{letter}{row + 1} = {val!r}", rows_affected=1)

    def _handle_range(self, rng: str):
        rs, cs, re, ce = _parse_range(rng)
        df = self._load()
        df = self._ensure_cols(df, ce)
        df = self._ensure_rows(df, re)
        sub = df.iloc[rs:re + 1, cs:ce + 1]
        rows = self._to_records(sub)
        return CommandResult(True, "RANGE", f"{rng}: {len(rows)} row(s) x {ce - cs + 1} col(s).", data=rows)

    def _handle_lastrow(self, col_letter: str):
        col = _col_letter_to_index(col_letter)
        df = self._load()
        if len(df) == 0 or col >= len(df.columns):
            return CommandResult(True, "LASTROW", f"Col {col_letter}: last row = 0 (empty)", value=0)
        ser = df.iloc[:, col]
        mask = ser.notna()
        if not mask.any():
            return CommandResult(True, "LASTROW", f"Col {col_letter}: last row = 0 (empty)", value=0)
        lr = mask[::-1].idxmax() + 1
        return CommandResult(True, "LASTROW", f"Col {col_letter}: last row = {lr}", value=lr)

    def _handle_lastcol(self, row_num: int):
        row = row_num - 1
        df = self._load()
        if row >= len(df) or len(df.columns) == 0:
            return CommandResult(True, "LASTCOL", f"Row {row_num}: last col = none", value=0)
        ser = df.iloc[row, :]
        mask = ser.notna()
        if not mask.any():
            return CommandResult(True, "LASTCOL", f"Row {row_num}: last col = none", value=0)
        col_name = mask[::-1].idxmax()
        for i, c in enumerate(df.columns):
            if c == col_name:
                return CommandResult(True, "LASTCOL",
                    f"Row {row_num}: last col = {_col_index_to_letter(i)}", value=i + 1)
        return CommandResult(True, "LASTCOL", f"Row {row_num}: last col = none", value=0)

    def _handle_insertrow(self, row_num: int):
        df = self._load()
        pos = min(row_num - 1, len(df))
        blank = pd.DataFrame([{c: None for c in (df.columns if len(df.columns) else ["A"])}])
        if len(df) == 0:
            df = blank
        else:
            df = pd.concat([df.iloc[:pos], blank, df.iloc[pos:]], ignore_index=True)
        self._save(df)
        return CommandResult(True, "INSERTROW", f"Inserted blank row at {row_num}.", rows_affected=1)

    def _handle_deleterow(self, row_num: int):
        df = self._load()
        pos = row_num - 1
        if pos < 0 or pos >= len(df):
            return CommandResult(False, "DELETEROW", f"Row {row_num} not found. {len(df)} row(s) exist.")
        df = df.drop(df.index[pos]).reset_index(drop=True)
        self._save(df)
        return CommandResult(True, "DELETEROW", f"Deleted row {row_num}.", rows_affected=1)

    def _handle_clearrow(self, row_num: int):
        df = self._load()
        pos = row_num - 1
        if pos < 0 or pos >= len(df):
            return CommandResult(False, "CLEARROW", f"Row {row_num} not found. {len(df)} row(s) exist.")
        df.iloc[pos, :] = None
        self._save(df)
        return CommandResult(True, "CLEARROW", f"Cleared row {row_num}.", rows_affected=1)

    def _handle_clear(self, rng: str):
        rs, cs, re, ce = _parse_range(rng)
        df = self._load()
        df = self._ensure_cols(df, ce)
        df = self._ensure_rows(df, re)
        df.iloc[rs:re + 1, cs:ce + 1] = None
        self._save(df)
        n = (re - rs + 1) * (ce - cs + 1)
        return CommandResult(True, "CLEAR", f"Cleared {n} cell(s) in {rng}.", rows_affected=n)

    def _handle_clearcol(self, col_letter: str):
        col = _col_letter_to_index(col_letter)
        df = self._load()
        if len(df.columns) == 0 or col >= len(df.columns):
            return CommandResult(False, "CLEARCOL", f"Column {col_letter} not found.")
        n = int(df.iloc[:, col].notna().sum())
        df.iloc[:, col] = None
        self._save(df)
        return CommandResult(True, "CLEARCOL", f"Cleared column {col_letter} ({n} cell(s)).", rows_affected=n)

    def _handle_addcol(self, name: str):
        df = self._load()
        if name in df.columns:
            return CommandResult(False, "ADDCOL", f"Column '{name}' already exists.")
        df[name] = None
        letter = _col_index_to_letter(len(df.columns) - 1)
        self._save(df)
        return CommandResult(True, "ADDCOL", f"Added '{name}' as column {letter}.")

    def _handle_rmvcol(self, name: str):
        df = self._load()
        if name not in df.columns:
            return CommandResult(False, "RMVCOL", f"Column '{name}' not found.")
        df.drop(columns=[name], inplace=True)
        self._save(df)
        return CommandResult(True, "RMVCOL", f"Deleted column '{name}'.")

    def _handle_rencol(self, old: str, new: str):
        df = self._load()
        if old not in df.columns:
            return CommandResult(False, "RENCOL", f"Column '{old}' not found.")
        if new in df.columns:
            return CommandResult(False, "RENCOL", f"Column '{new}' already exists.")
        df.rename(columns={old: new}, inplace=True)
        self._save(df)
        return CommandResult(True, "RENCOL", f"Renamed '{old}' → '{new}'.")

    def _handle_clr(self):
        df = self._load()
        n = len(df)
        if n == 0:
            return CommandResult(True, "CLR", "Already empty.")
        df = df.iloc[0:0]
        self._save(df)
        return CommandResult(True, "CLR", f"Cleared {n} row(s). Columns kept.")
