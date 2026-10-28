"""Orchestrate one run: read -> map -> build -> de-duplicate -> reconcile -> write (design F.2).

UI-agnostic: the CLI and the Streamlit app both call `run()`.
"""

from __future__ import annotations

import logging
import uuid
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from . import audit
from .build import RecordBuilder
from .config import Config, file_sha256, load_config
from .history import BatchRow, Ledger
from .mapping import (ColumnMap, decisions_from_preview, header_signature, load_decisions, resolve,
                      save_decisions)
from .model import Issue, Record, RecordStatus, Severity, SourceResponse
from .reader import SheetData, choose_sheet, list_sheets, read_sheet
from .writer import Layout, leads_log_row, read_template_header, write_leads_log, write_masterlist

log = logging.getLogger("fuel_parser")


class NeedsConfirmation(Exception):
    """The run cannot continue until the user confirms the sheet or mapping items."""

    def __init__(self, message: str, cmap: ColumnMap | None = None, sheet: str | None = None, file: str | None = None):
        super().__init__(message)
        self.cmap = cmap
        self.sheet = sheet
        self.file = file


class ReconciliationError(RuntimeError):
    """Counts do not balance; nothing is exported (AC2)."""


@dataclass
class RunOptions:
    source: Path | list[Path]  # one or several response exports (bulk import, H-35)
    template: Path | None  # None = built-in MASTERLIST layout (H-34)
    out_dir: Path
    config_dir: Path = Path("config")
    ledger_path: Path = Path("output/import_history.sqlite")
    profiles_dir: Path = Path("profiles")
    sellers_path: Path | None = None
    teams_path: Path | None = None
    sheet: str | None = None
    last_id: int | None = None
    team_filter: list[str] | None = None
    leads_log: bool = False
    dry_run: bool = False
    timestamp_formats: list[str] | None = None
    accept_suggestions: bool = False
    ignore_unmapped: bool = False
    team_sheet_check: bool = False  # opt-in: other sheets are ignored by default (H-33)
    now: datetime | None = None

    @property
    def sources(self) -> list[Path]:
        items = self.source if isinstance(self.source, (list, tuple)) else [self.source]
        return [Path(s) for s in items]


@dataclass
class SourceFile:
    """One response export in the run, with its own sheet, mapping and checksum."""

    path: Path
    sheet: SheetData
    cmap: ColumnMap
    sha256: str

    @property
    def name(self) -> str:
        return self.path.name


@dataclass
class RunResult:
    batch_id: str
    sources: list[SourceFile]
    records: list[Record]
    response_issues: list[Issue]
    counts: dict[str, int]
    checks: list[tuple[str, bool, str]]
    previous_batches: list[BatchRow]
    layout: Layout
    outputs: dict[str, Path] = field(default_factory=dict)
    cfg: Config | None = None
    dry_run: bool = False

    @property
    def exported(self) -> list[Record]:
        return [r for r in self.records if r.status == RecordStatus.EXPORTED]

    # The first source, for callers that handle a single file.
    @property
    def sheet(self) -> SheetData:
        return self.sources[0].sheet

    @property
    def cmap(self) -> ColumnMap:
        return self.sources[0].cmap


def _nonempty_counts(sd: SheetData) -> dict[int, int]:
    counts: dict[int, int] = defaultdict(int)
    for _, row in sd.rows:
        for i, v in enumerate(row):
            if v is not None and not (isinstance(v, str) and v.strip() == ""):
                counts[i] += 1
    return counts


def prepare_source(cfg: Config, path: Path, opts: RunOptions) -> tuple[SheetData, ColumnMap]:
    """Choose the response sheet of one file and resolve its mapping."""
    src = cfg.profile.source
    label = f"{path.name}: " if len(opts.sources) > 1 else ""
    if opts.sheet and opts.sheet in list_sheets(path):
        sheet = opts.sheet
        ignored = [n for n in list_sheets(path) if n != sheet]
    else:
        choice = choose_sheet(path, src["preferred_sheet"], src["signature_keys"], src.get("header_row", 1))
        sheet, ignored = choice.name, choice.ignored
        if choice.needs_confirmation:
            raise NeedsConfirmation(f"{label}several sheets look like response sheets. Proposed sheet: '{sheet}'. "
                                    "Confirm it by passing the sheet name explicitly.", sheet=sheet, file=path.name)
    sd = read_sheet(path, sheet, src.get("header_row", 1))
    sd.ignored_sheets = ignored
    nonempty = _nonempty_counts(sd)
    sig = header_signature(sd.headers)
    decisions = load_decisions(opts.profiles_dir, sig)
    cmap = resolve(sd.headers, cfg.profile, nonempty, decisions)
    if cmap.needs_confirmation and (opts.accept_suggestions or opts.ignore_unmapped):
        new = decisions_from_preview(cmap, sd.headers, opts.accept_suggestions, opts.ignore_unmapped)
        if new:
            save_decisions(opts.profiles_dir, sig, new)
            decisions.update(new)
            cmap = resolve(sd.headers, cfg.profile, nonempty, decisions)
    if cmap.needs_confirmation:
        raise NeedsConfirmation(f"{label}{len(cmap.needs_confirmation)} mapping item(s) need confirmation",
                                cmap=cmap, sheet=sheet, file=path.name)
    return sd, cmap


def prepare(opts: RunOptions) -> tuple[Config, list[SourceFile], Layout]:
    """Load config, resolve the output layout, then choose the sheet and mapping of every source file."""
    cfg = load_config(opts.config_dir, opts.sellers_path, opts.teams_path)
    layout = read_template_header(opts.template, cfg.profile)  # stops on a wrong template before any output
    if not opts.sources:
        raise ValueError("Add at least one response export.")
    names = [p.name for p in opts.sources]
    dupes = sorted({n for n in names if names.count(n) > 1})
    shas = [file_sha256(p) for p in opts.sources]
    if len(set(shas)) != len(shas):
        raise ValueError("The same response export was added twice. Remove the duplicate file.")
    if dupes:
        raise ValueError(f"Two response exports share the name {', '.join(dupes)}. Rename one so the audit can tell them apart.")
    files = []
    for path, sha in zip(opts.sources, shas):
        sd, cmap = prepare_source(cfg, path, opts)
        files.append(SourceFile(path, sd, cmap, sha))
    return cfg, files, layout


def _dedupe(records: list[Record], cfg: Config, ledger: Ledger) -> None:
    imported = ledger.imported_keys()
    ledger_hashes = ledger.imported_content_hashes()
    seen_hash: dict[str, Record] = {}
    seen_key: dict[str, Record] = {}
    ordered = sorted(records, key=lambda r: (r.timestamp or datetime.max, r.source_file, r.source_row, r.block_index or 0))
    where = lambda r: f"{r.source_file} row {r.source_row}" if r.source_file else f"source row {r.source_row}"
    for rec in ordered:
        if rec.src_key in imported:
            rec.status = RecordStatus.ALREADY_IMPORTED
            continue
        same_key = seen_key.get(rec.src_key)
        first = seen_hash.get(rec.content_hash)
        if same_key is not None:  # the same response appears in more than one export of this run
            rec.add("PROBABLE_DUPLICATE", cfg.policy.sev("PROBABLE_DUPLICATE"), "", f"Same response as {where(same_key)}")
        elif first is not None:
            rec.add("PROBABLE_DUPLICATE", cfg.policy.sev("PROBABLE_DUPLICATE"), "", f"Same content as {where(first)}")
        elif rec.content_hash in ledger_hashes:
            rec.add("PROBABLE_DUPLICATE", cfg.policy.sev("PROBABLE_DUPLICATE"), "",
                    "Same content as a record exported in an earlier batch")
        else:
            seen_hash[rec.content_hash] = rec
            seen_key[rec.src_key] = rec
    groups: dict[str, list[Record]] = defaultdict(list)
    for rec in records:
        if rec.client_key:
            groups[rec.client_key].append(rec)
    n = 0
    for members in groups.values():
        if len({m.src_key for m in members}) < 2:
            continue  # one response repeated across files is a duplicate, not a second client registration
        n += 1
        for rec in members:
            rec.duplicate_group = f"G{n:04d}"
            rec.add("POSSIBLE_DUPLICATE_CLIENT", cfg.policy.sev("POSSIBLE_DUPLICATE_CLIENT"), "",
                    f"Same client name and mobile in {len(members)} records (group {rec.duplicate_group})")


def _assign_status(rec: Record, team_filter: list[str] | None) -> None:
    if rec.status == RecordStatus.ALREADY_IMPORTED:
        return
    if team_filter and not any(rec.sales_leader.upper().startswith(t.upper()) for t in team_filter):
        rec.status = RecordStatus.EXCLUDED_BY_FILTER
    elif rec.has(Severity.ERROR):
        rec.status = RecordStatus.REVIEW
    elif rec.has(Severity.HELD):
        rec.status = RecordStatus.DUPLICATE_HELD
    else:
        rec.status = RecordStatus.EXPORTED


def reconcile(sheets: list[SheetData] | SheetData, records: list[Record], response_issues: list[Issue]) -> tuple[dict, list]:
    sheets = sheets if isinstance(sheets, list) else [sheets]
    by_status = Counter(r.status.value for r in records)
    per_row = Counter((r.source_file, r.source_row) for r in records)
    no_client_rows = {(i.source_file, i.source_row) for i in response_issues if i.code == "NO_CLIENT_DATA"}
    extra_blocks = sum(n - 1 for n in per_row.values() if n > 1)
    responses = sum(len(sd.rows) for sd in sheets)
    blanks = sum(sd.blank_rows for sd in sheets)
    counts = {
        "source_files": len(sheets),
        "data_rows_read": responses + blanks,
        "blank_rows": blanks,
        "responses": responses,
        "responses_no_client": len(no_client_rows),
        "responses_multi_block": sum(1 for n in per_row.values() if n > 1),
        "records_built": len(records),
        **{f"records_{s.value.lower()}": by_status.get(s.value, 0) for s in RecordStatus},
    }
    outcome = sum(by_status.values())
    checks = [
        ("data rows = blank rows + responses", counts["data_rows_read"] == blanks + responses, ""),
        ("records built = responses − no-client + extra blocks",
         counts["records_built"] == responses - len(no_client_rows) + extra_blocks, ""),
        ("records built = exported + review + duplicate held + already imported + excluded",
         counts["records_built"] == outcome and all(r.status is not None for r in records), ""),
        ("every response accounted for exactly once (records or NO_CLIENT_DATA)",
         len(per_row) + len(no_client_rows) == responses and not (set(per_row) & no_client_rows), ""),
    ]
    return counts, checks


def run(opts: RunOptions) -> RunResult:
    cfg, files, layout = prepare(opts)
    out_dir = Path(opts.out_dir)
    ledger = Ledger(opts.ledger_path)
    previous: dict[str, BatchRow] = {}
    for f in files:
        for b in ledger.batches_for_file(f.sha256):
            previous[b.batch_id] = b

    records: list[Record] = []
    response_issues: list[Issue] = []
    multi = len(files) > 1
    for f in files:
        builder = RecordBuilder(cfg, f.cmap, opts.timestamp_formats, opts.now)
        for row_no, values in f.sheet.rows:
            recs, issues = builder.build(SourceResponse(row_no, dict(enumerate(values))))
            for rec in recs:
                rec.source_file = f.name if multi else ""
            for issue in issues:
                issue.source_file = f.name if multi else ""
            records.extend(recs)
            response_issues.extend(issues)

    _dedupe(records, cfg, ledger)
    for rec in records:
        _assign_status(rec, opts.team_filter)

    counts, checks = reconcile([f.sheet for f in files], records, response_issues)
    if not all(ok for _, ok, _ in checks):
        failed = [name for name, ok, _ in checks if not ok]
        raise ReconciliationError("Reconciliation failed; nothing exported: " + "; ".join(failed))

    batch_id = datetime.now().strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
    result = RunResult(batch_id, files, records, response_issues, counts, checks, list(previous.values()), layout,
                       cfg=cfg, dry_run=opts.dry_run)
    exported = sorted(result.exported, key=lambda r: (r.source_file, r.source_row, r.block_index or 0))
    if exported and not opts.dry_run:
        if opts.last_id is None:
            raise ValueError(f"{len(exported)} record(s) are ready to export: enter the last Id in the live MASTERLIST (--last-id).")
        for i, rec in enumerate(exported, start=opts.last_id + 1):
            rec.dest_id = i

    out_dir.mkdir(parents=True, exist_ok=True)
    stem = files[0].path.stem.replace(" ", "_")[:40] + (f"_+{len(files) - 1}" if multi else "")
    inputs = {f.path.resolve() for f in files} | ({Path(opts.template).resolve()} if opts.template else set())

    def out_path(prefix: str) -> Path:
        p = out_dir / f"{prefix}_{stem}_{batch_id}.xlsx"
        if p.resolve() in inputs:
            raise ValueError("Output path would overwrite an input file")
        return p

    try:
        if not opts.dry_run:
            p = out_path("FUEL_MASTERLIST_IMPORT")
            result.outputs["import"] = p
            for rec in write_masterlist(p, layout, exported, cfg.profile):
                rec.add("FORMULA_TEXT_ESCAPED", cfg.policy.sev("FORMULA_TEXT_ESCAPED"), "", "Text starting with '=' written as literal text")
        if opts.leads_log:
            p = out_path("FUEL_LEADS_LOG")
            result.outputs["leads_log"] = p
            write_leads_log(p, [leads_log_row(r.values, r.seller_slot, cfg.profile) for r in exported], cfg.profile)
        p = out_path("FUEL_IMPORT_AUDIT")
        result.outputs["audit"] = p
        audit.write_audit(p, result, opts)
        if not opts.dry_run:
            ledger.record_export({
                "batch_id": batch_id, "source_file_name": "; ".join(f.name for f in files),
                "source_file_sha256": ";".join(f.sha256 for f in files),
                "source_sheet": "; ".join(f.sheet.sheet_name for f in files),
                "mapping_profile_id": cfg.profile.raw.get("profile_id", ""),
                "responses_read": counts["responses"], "records_built": counts["records_built"],
                "records_exported": len(exported), "records_review": counts["records_review"],
                "records_skipped": counts["records_already_imported"], "output_path": str(result.outputs["import"]),
            }, [(r.src_key, r.source_row, r.dest_id, r.content_hash) for r in exported])
    except Exception:
        # Ledger keys are written last, in one transaction; on any failure no output of this run is left behind.
        for p in result.outputs.values():
            Path(p).unlink(missing_ok=True)
        result.outputs.clear()
        raise
    log.info("batch %s: %s", batch_id, {k: v for k, v in counts.items() if k.startswith("records_")})
    return result
