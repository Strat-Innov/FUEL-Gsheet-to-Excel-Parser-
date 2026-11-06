"""Write the import-ready MASTERLIST and the optional Leads Log (design D.6, F.5, I.2)."""

from __future__ import annotations

import functools
from copy import copy
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


_SLOT_FIELDS = {n: (f"seller_name_{n}", f"seller_email_{n}") for n in (1, 2, 3)}


@dataclass(frozen=True)
class Layout:
    """Destination column layout: exact header cells, the field each column holds (None = unrecognised), widths."""

    headers: tuple
    fields: tuple
    widths: tuple  # (column letter, width) pairs
    origin: str  # "built-in" or "<file name> · sheet <name>"
    reordered: bool = False
    sheet_title: str = "MASTERLIST"
    missing: tuple = ()  # built-in fields this template has no column for (not exported)
    unknown: tuple = ()  # template headers with no known meaning (written blank)
    header_styles: tuple = ()  # per column: copied header cell style, or empty

    @property
    def slots(self) -> tuple:
        """Seller pairs (1, 2, 3) for which the layout has both the name and the e-mail column."""
        return tuple(n for n, pair in _SLOT_FIELDS.items() if all(f in self.fields for f in pair))

    def slot_letters(self, n: int) -> str:
        return "/".join(get_column_letter(self.fields.index(f) + 1) for f in _SLOT_FIELDS[n])


def builtin_layout(profile: Profile) -> Layout:
    """The fixed 35-column MASTERLIST layout carried in the mapping profile (H-34)."""
    tpl = profile.dest["template"]
    widths = tuple((get_column_letter(i), t.get("width")) for i, t in enumerate(tpl, start=1))
    return Layout(tuple(t["header"] for t in tpl), tuple(t["field"] for t in tpl), widths, "built-in",
                  sheet_title=profile.dest["sheet"])


def _style_of(cell) -> dict:
    return {"font": copy(cell.font), "fill": copy(cell.fill), "alignment": copy(cell.alignment),
            "border": copy(cell.border)} if cell.has_style else {}


@functools.lru_cache(maxsize=4)
def _layout_from_file(path: str, size: int, mtime: float, preferred: str, key_to_field: tuple, ref_fields: tuple,
                      required: tuple, ref_headers: tuple) -> Layout:
    k2f = dict(key_to_field)
    label = {f: " ".join(str(h).split()) for f, h in zip(ref_fields, ref_headers)}  # readable header names for messages
    wb = openpyxl.load_workbook(path, read_only=False)
    candidates, problems = [], []
    try:
        names = [preferred] * (preferred in wb.sheetnames) + [n for n in wb.sheetnames if n != preferred]
        for name in names:
            ws = wb[name]
            header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            while header and header[-1] is None:
                header.pop()
            fields = [k2f.get(match_key(h)) if h is not None else None for h in header]
            known = [f for f in fields if f]
            if not known:
                continue
            dupes = sorted({f for f in known if known.count(f) > 1})
            lacking = [f for f in required if f not in known]
            has_pair = any(all(f in known for f in pair) for pair in _SLOT_FIELDS.values())
            if dupes or lacking or not has_pair:
                problems.append((len(known), name, dupes, lacking, has_pair))
                continue
            widths = tuple((get_column_letter(i), ws.column_dimensions[get_column_letter(i)].width)
                           for i in range(1, len(header) + 1))
            layout = Layout(
                tuple(header), tuple(fields), widths, f"{Path(path).name} · sheet {name}",
                reordered=tuple(f for f in fields if f) != tuple(f for f in ref_fields if f in known),
                sheet_title=name,
                missing=tuple(f for f in ref_fields if f not in known),
                unknown=tuple(str(h).strip() for h, f in zip(header, fields) if f is None and h is not None),
                header_styles=tuple(_style_of(ws.cell(1, c)) for c in range(1, len(header) + 1)))
            candidates.append((len(known), layout))
    finally:
        wb.close()
    if candidates:
        return max(candidates, key=lambda c: c[0])[1]  # the sheet with the most recognised MASTERLIST columns
    detail = ""
    if problems:
        _, name, dupes, lacking, has_pair = max(problems)
        reasons = ([f"missing {', '.join(label[f] for f in lacking)}"] if lacking else []) + \
                  ([f"repeated columns {', '.join(label[f] for f in dupes)}"] if dupes else []) + \
                  ([] if has_pair else ["no complete Seller Name / Seller's Email Address pair"])
        detail = f" Closest sheet '{name}': {'; '.join(reasons)}."
    raise TemplateError("The template file has no usable MASTERLIST header row." + detail +
                        " Leave the template empty to use the built-in layout.")


def read_template_header(template: Path | None, profile: Profile) -> Layout:
    """Return the output layout: built-in when no template is given, otherwise the template's own format.

    The output follows the supplied template (H-36): its columns and order, its sheet name and header style.
    Columns are matched to fields by header meaning; built-in fields the template lacks are not exported and
    unrecognised template columns are left blank (both reported in the audit).
    """
    base = builtin_layout(profile)
    if template is None:
        return base
    key_to_field = tuple((match_key(h), f) for h, f in zip(base.headers, base.fields))
    st = Path(template).stat()
    return _layout_from_file(str(Path(template).resolve()), st.st_size, st.st_mtime, profile.dest["sheet"],
                             key_to_field, base.fields, tuple(profile.dest.get("template_required") or ()), base.headers)


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
    ws.title = layout.sheet_title  # the template's own sheet name (H-36)
    header = layout.headers
    for c, h in enumerate(header, start=1):
        cell = ws.cell(1, c)
        cell.value = h  # byte-exact header (AC8)
        style = layout.header_styles[c - 1] if c - 1 < len(layout.header_styles) else {}
        for attr, value in style.items():
            setattr(cell, attr, value)
    for letter, w in layout.widths:
        if w:
            ws.column_dimensions[letter].width = w
    dt_cols = set(profile.dest.get("datetime_columns") or [])
    dt_fmt = profile.dest.get("datetime_format", "m/d/yy h:mm")
    escaped = []
    for r, rec in enumerate(records, start=2):
        for c, fid in enumerate(layout.fields, start=1):
            val = rec.values.get(fid) if fid else None  # unrecognised template column: left blank
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
