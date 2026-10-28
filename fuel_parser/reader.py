"""Read a static export (.xlsx or .csv) into headers and rows (design M.1, M.2).

Input files are only ever opened for reading.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from .headers import match_key


class ReadError(ValueError):
    """Raised when the input cannot be used as a response export."""


@dataclass
class SheetData:
    path: Path
    sheet_name: str
    headers: list[object]
    rows: list[tuple[int, list[object]]]  # (1-based source row number, values)
    blank_rows: int
    is_csv: bool
    ignored_sheets: list[str] = field(default_factory=list)


def list_sheets(path: Path) -> list[str]:
    if path.suffix.lower() == ".csv":
        return [path.stem]
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _signature_ok(headers: list[object], signature_keys: list[str]) -> bool:
    keys = {match_key(h) for h in headers}
    return all(k in keys for k in signature_keys)


@dataclass
class SheetChoice:
    name: str
    needs_confirmation: bool
    ignored: list[str]  # every other sheet in the workbook; never read for data


def choose_sheet(path: Path, preferred: str, signature_keys: list[str], header_row: int = 1) -> SheetChoice:
    """Pick the one response sheet and ignore all others.

    The response sheet is recognised by its header row (Timestamp, SALES LEADER, Lead Source and the
    Seller Name client blocks). The preferred name (MARF RAW) wins when it carries those headers.
    Otherwise a single matching sheet is used automatically; only several matching sheets need the
    user to choose. Derived team sheets never carry the signature, so they are never chosen (T39).
    """
    if path.suffix.lower() == ".csv":
        return SheetChoice(path.stem, False, [])
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        names = list(wb.sheetnames)
        matching = []
        for name in names:
            first = next(wb[name].iter_rows(min_row=header_row, max_row=header_row, values_only=True), ())
            if _signature_ok(list(first), signature_keys):
                matching.append(name)
    finally:
        wb.close()
    others = lambda chosen: [n for n in names if n != chosen]
    if preferred in matching or (preferred in names and not matching):
        return SheetChoice(preferred, False, others(preferred))
    if len(matching) == 1:
        return SheetChoice(matching[0], False, others(matching[0]))
    if matching:
        return SheetChoice(matching[0], True, others(matching[0]))
    raise ReadError(
        "No response sheet found: no sheet has the Timestamp / SALES LEADER / Lead Source / Seller Name headers. "
        "Derived team sheets cannot be used as input."
    )


def _is_blank(v: object) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def read_sheet(path: Path, sheet_name: str | None, header_row: int = 1) -> SheetData:
    path = Path(path)
    if not path.exists():
        raise ReadError(f"File not found: {path}")
    if path.suffix.lower() == ".csv":
        with open(path, encoding="utf-8-sig", newline="") as fh:
            all_rows = list(csv.reader(fh))
        is_csv = True
        name = path.stem
    elif path.suffix.lower() in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            if sheet_name not in wb.sheetnames:
                raise ReadError(f"Sheet '{sheet_name}' not found in {path.name}")
            all_rows = [list(r) for r in wb[sheet_name].iter_rows(values_only=True)]
        finally:
            wb.close()
        is_csv = False
        name = sheet_name
    else:
        raise ReadError(f"Unsupported file type: {path.suffix} (use .xlsx or .csv)")
    if len(all_rows) < header_row:
        raise ReadError("The sheet has no header row")
    headers = list(all_rows[header_row - 1])
    width = len(headers)
    rows, blanks = [], 0
    for n, r in enumerate(all_rows[header_row:], start=header_row + 1):
        r = list(r) + [None] * (width - len(r))
        if all(_is_blank(v) for v in r):
            blanks += 1
            continue
        rows.append((n, r[:width]))
    # Trailing blank rows exported by spreadsheet tools are not data rows.
    return SheetData(path, name, headers, rows, blanks, is_csv)
