"""Write the import-ready MASTERLIST and the optional Leads Log (design D.6, F.5, I.2)."""

from __future__ import annotations

import functools
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from .config import Profile
from .headers import match_key
from .model import Record


class TemplateError(ValueError):
    """Raised when the destination template does not match the approved schema."""


@dataclass(frozen=True)
class Layout:
    """Destination column layout: exact header cells, the field each column holds, and widths."""

    headers: tuple
    fields: tuple
    widths: tuple  # (column letter, width) pairs
    origin: str  # "built-in" or "<file name> · sheet <name>"
    reordered: bool = False


def builtin_layout(profile: Profile) -> Layout:
    """The fixed 35-column MASTERLIST layout carried in the mapping profile (H-34)."""
    tpl = profile.dest["template"]
    widths = tuple((get_column_letter(i), t.get("width")) for i, t in enumerate(tpl, start=1))
    return Layout(tuple(t["header"] for t in tpl), tuple(t["field"] for t in tpl), widths, "built-in")


@functools.lru_cache(maxsize=4)
def _layout_from_file(path: str, size: int, mtime: float, preferred: str, key_to_field: tuple, ref_fields: tuple) -> Layout:
    k2f = dict(key_to_field)
    wb = openpyxl.load_workbook(path, read_only=False)
    best = None  # closest non-matching sheet, for a useful error message
    try:
        names = [preferred] * (preferred in wb.sheetnames) + [n for n in wb.sheetnames if n != preferred]
        for name in names:
            ws = wb[name]
            header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            while header and header[-1] is None:
                header.pop()
            keys = [match_key(h) for h in header]
            fields = [k2f.get(k) for k in keys]
            if len(header) == len(k2f) and None not in fields and len(set(fields)) == len(fields):
                widths = tuple((get_column_letter(i), ws.column_dimensions[get_column_letter(i)].width)
                               for i in range(1, len(header) + 1))
                return Layout(tuple(header), tuple(fields), widths, f"{Path(path).name} · sheet {name}",
                              tuple(fields) != ref_fields)
            overlap = sum(1 for f in fields if f)
            if overlap and (best is None or overlap > best[0]):
                missing = [k for k in k2f if k not in keys]
                unknown = [str(h).strip() for h, f in zip(header, fields) if not f and h is not None]
                best = (overlap, name, missing, unknown)
    finally:
        wb.close()
    detail = ""
    if best:
        _, name, missing, unknown = best
        detail = (f" Closest sheet '{name}': missing {len(missing)} header(s)"
                  + (f" ({', '.join(missing[:5])}{'…' if len(missing) > 5 else ''})" if missing else "")
                  + (f"; unrecognised: {', '.join(u[:40] for u in unknown[:5])}" if unknown else "") + ".")
    raise TemplateError("No sheet in the template file has the 35 MASTERLIST headers (in any order). "
                        "No output was written." + detail + " Leave the template empty to use the built-in layout.")


def read_template_header(template: Path | None, profile: Profile) -> Layout:
    """Return the output layout: built-in when no template is given, otherwise the template's own order.

    A supplied template must contain exactly the 35 approved headers. They are matched by meaning, so a
    re-arranged template is accepted and the output follows its column order (H-34).
    """
    base = builtin_layout(profile)
    if template is None:
        return base
    key_to_field = tuple((match_key(h), f) for h, f in zip(base.headers, base.fields))
    st = Path(template).stat()
    return _layout_from_file(str(Path(template).resolve()), st.st_size, st.st_mtime, profile.dest["sheet"],
                             key_to_field, base.fields)


def set_text(cell, value: object) -> bool:
    """Write a value as a literal string. Returns True when formula-like text was escaped."""
    cell.value = value
    cell.number_format = "@"
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"  # openpyxl would otherwise store it as a formula (D.6)
        return True
    return False


def write_masterlist(path: Path, layout: Layout, records: list[Record], profile: Profile) -> list[Record]:
    """Write exported records in the layout's column order. Returns records whose text needed formula escaping."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = profile.dest["sheet"]
    header = layout.headers
    for c, h in enumerate(header, start=1):
        ws.cell(1, c).value = h  # byte-exact header (AC8)
    for letter, w in layout.widths:
        if w:
            ws.column_dimensions[letter].width = w
    dt_cols = set(profile.dest.get("datetime_columns") or [])
    dt_fmt = profile.dest.get("datetime_format", "m/d/yy h:mm")
    escaped = []
    for r, rec in enumerate(records, start=2):
        for c, fid in enumerate(layout.fields, start=1):
            val = rec.values.get(fid)
            cell = ws.cell(r, c)
            if fid == "id":
                cell.value = rec.dest_id
            elif fid in dt_cols:
                cell.value = val
                cell.number_format = dt_fmt
            elif val is None:
                cell.number_format = "@"
            elif set_text(cell, str(val)):
                escaped.append(rec)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{max(1, len(records) + 1)}"
    wb.save(path)
    return escaped


def leads_log_row(values: dict, slot: int | None, profile: Profile) -> list[object]:
    """Project one MASTERLIST record onto the 17 Leads Log columns (I.2)."""
    ls = values.get("lead_source")
    detail = profile.lead_source_detail.get(ls) if ls else None
    specific = values.get(detail) if detail and detail != "question" else None
    return [
        values.get("completion_time"),
        values.get(f"seller_email_{slot}") if slot else None,
        values.get("sales_leader"),
        values.get(f"seller_name_{slot}") if slot else None,
        values.get("surname"), values.get("first_name"), values.get("middle_name"),
        None if values.get("mobile") is None else str(values.get("mobile")),
        values.get("client_email"), values.get("residence"),
        values.get("project_first"), values.get("project_second"), values.get("remarks"),
        ls, specific,
        None,  # LEAD SOURCE (DISTINCT): blank until a rule is approved (H-23)
        values.get("awareness"),
    ]


def slot_mismatch(values: dict, slot: int) -> bool:
    """True when the configured slot is empty while another slot holds seller data (N.2)."""
    def filled(n: int) -> bool:
        return bool(values.get(f"seller_name_{n}") or values.get(f"seller_email_{n}"))
    return not filled(slot) and any(filled(n) for n in (1, 2, 3) if n != slot)


def write_leads_log(path: Path, rows: list[list[object]], profile: Profile) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "LEADS LOG"
    headers = profile.leads_log["headers"]
    for c, h in enumerate(headers, start=1):
        ws.cell(1, c).value = h
    fmt = profile.leads_log.get("timestamp_format", "mmm/dd/yyyy hh:mm:ss")
    for r, row in enumerate(rows, start=2):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(r, c)
            if c == 1 and isinstance(val, datetime):
                cell.value = val
                cell.number_format = fmt  # H-22: real datetime
            elif val is None:
                cell.number_format = "@"
            else:
                set_text(cell, str(val))
    wb.save(path)
