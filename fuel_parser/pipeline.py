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
from .writer import leads_log_row, read_template_header, write_leads_log, write_masterlist

log = logging.getLogger("fuel_parser")


class NeedsConfirmation(Exception):
    """The run cannot continue until the user confirms the sheet or mapping items."""

    def __init__(self, message: str, cmap: ColumnMap | None = None, sheet: str | None = None):
        super().__init__(message)
        self.cmap = cmap
        self.sheet = sheet


class ReconciliationError(RuntimeError):
    """Counts do not balance; nothing is exported (AC2)."""


@dataclass
class RunOptions:
    source: Path
    template: Path
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


@dataclass
class RunResult:
    batch_id: str
    sheet: SheetData
    cmap: ColumnMap
    records: list[Record]
    response_issues: list[Issue]
    counts: dict[str, int]
    checks: list[tuple[str, bool, str]]
    previous_batches: list[BatchRow]
    source_sha256: str
    outputs: dict[str, Path] = field(default_factory=dict)
    cfg: Config | None = None
    dry_run: bool = False

    @property
    def exported(self) -> list[Record]:
        return [r for r in self.records if r.status == RecordStatus.EXPORTED]


def _nonempty_counts(sd: SheetData) -> dict[int, int]:
    counts: dict[int, int] = defaultdict(int)
    for _, row in sd.rows:
        for i, v in enumerate(row):
            if v is not None and not (isinstance(v, str) and v.strip() == ""):
                counts[i] += 1
    return counts


def prepare(opts: RunOptions) -> tuple[Config, SheetData, ColumnMap]:
    """Load config, validate the template, choose the sheet and resolve the mapping."""
    cfg = load_config(opts.config_dir, opts.sellers_path, opts.teams_path)
    read_template_header(opts.template, cfg.profile)  # stops on schema mismatch before any output
    src = cfg.profile.source
    if opts.sheet:
        sheet = opts.sheet
        ignored = [n for n in list_sheets(opts.source) if n != sheet]
    else:
        choice = choose_sheet(opts.source, src["preferred_sheet"], src["signature_keys"], src.get("header_row", 1))
        sheet, ignored = choice.name, choice.ignored
        if choice.needs_confirmation:
            raise NeedsConfirmation(f"Several sheets look like response sheets. Proposed sheet: '{sheet}'. "
                                    "Confirm it by passing the sheet name explicitly.", sheet=sheet)
    sd = read_sheet(opts.source, sheet, src.get("header_row", 1))
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
        raise NeedsConfirmation(f"{len(cmap.needs_confirmation)} mapping item(s) need confirmation", cmap=cmap, sheet=sheet)
    return cfg, sd, cmap


def _dedupe(records: list[Record], cfg: Config, ledger: Ledger) -> None:
    imported = ledger.imported_keys()
    ledger_hashes = ledger.imported_content_hashes()
    seen_hash: dict[str, Record] = {}
    ordered = sorted(records, key=lambda r: (r.timestamp or datetime.max, r.source_row, r.block_index or 0))
    for rec in ordered:
        if rec.src_key in imported:
            rec.status = RecordStatus.ALREADY_IMPORTED
            continue
        first = seen_hash.get(rec.content_hash)
        if first is not None:
            rec.add("PROBABLE_DUPLICATE", cfg.policy.sev("PROBABLE_DUPLICATE"), "",
                    f"Same content as source row {first.source_row}")
        elif rec.content_hash in ledger_hashes:
            rec.add("PROBABLE_DUPLICATE", cfg.policy.sev("PROBABLE_DUPLICATE"), "",
                    "Same content as a record exported in an earlier batch")
        else:
            seen_hash[rec.content_hash] = rec
    groups: dict[str, list[Record]] = defaultdict(list)
    for rec in records:
        if rec.client_key:
            groups[rec.client_key].append(rec)
    n = 0
    for members in groups.values():
        if len(members) < 2:
            continue
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


def reconcile(sd: SheetData, records: list[Record], response_issues: list[Issue]) -> tuple[dict, list]:
    by_status = Counter(r.status.value for r in records)
    per_row = Counter(r.source_row for r in records)
    no_client = len({i.source_row for i in response_issues if i.code == "NO_CLIENT_DATA"})
    extra_blocks = sum(n - 1 for n in per_row.values() if n > 1)
    counts = {
        "data_rows_read": len(sd.rows) + sd.blank_rows,
        "blank_rows": sd.blank_rows,
        "responses": len(sd.rows),
        "responses_no_client": no_client,
        "responses_multi_block": sum(1 for n in per_row.values() if n > 1),
        "records_built": len(records),
        **{f"records_{s.value.lower()}": by_status.get(s.value, 0) for s in RecordStatus},
    }
    outcome = sum(by_status.values())
    responses_accounted = len(per_row) + no_client
    checks = [
        ("data rows = blank rows + responses", counts["data_rows_read"] == counts["blank_rows"] + counts["responses"], ""),
        ("records built = responses − no-client + extra blocks",
         counts["records_built"] == counts["responses"] - no_client + extra_blocks, ""),
        ("records built = exported + review + duplicate held + already imported + excluded",
         counts["records_built"] == outcome and all(r.status is not None for r in records), ""),
        ("every response accounted for exactly once (records or NO_CLIENT_DATA)",
         responses_accounted == counts["responses"] and not (set(per_row) & {i.source_row for i in response_issues}), ""),
    ]
    return counts, checks


def run(opts: RunOptions) -> RunResult:
    cfg, sd, cmap = prepare(opts)
    out_dir = Path(opts.out_dir)
    ledger = Ledger(opts.ledger_path)
    source_sha = file_sha256(opts.source)
    previous = ledger.batches_for_file(source_sha)

    builder = RecordBuilder(cfg, cmap, opts.timestamp_formats, opts.now)
    records: list[Record] = []
    response_issues: list[Issue] = []
    for row_no, values in sd.rows:
        recs, issues = builder.build(SourceResponse(row_no, dict(enumerate(values))))
        records.extend(recs)
        response_issues.extend(issues)

    _dedupe(records, cfg, ledger)
    for rec in records:
        _assign_status(rec, opts.team_filter)

    counts, checks = reconcile(sd, records, response_issues)
    if not all(ok for _, ok, _ in checks):
        failed = [name for name, ok, _ in checks if not ok]
        raise ReconciliationError("Reconciliation failed; nothing exported: " + "; ".join(failed))

    batch_id = datetime.now().strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
    result = RunResult(batch_id, sd, cmap, records, response_issues, counts, checks, previous, source_sha,
                       cfg=cfg, dry_run=opts.dry_run)
    exported = sorted(result.exported, key=lambda r: (r.source_row, r.block_index or 0))
    if exported and not opts.dry_run:
        if opts.last_id is None:
            raise ValueError(f"{len(exported)} record(s) are ready to export: enter the last Id in the live MASTERLIST (--last-id).")
        for i, rec in enumerate(exported, start=opts.last_id + 1):
            rec.dest_id = i

    out_dir.mkdir(parents=True, exist_ok=True)
    stem = Path(opts.source).stem.replace(" ", "_")[:40]
    inputs = {Path(opts.source).resolve(), Path(opts.template).resolve()}

    def out_path(prefix: str) -> Path:
        p = out_dir / f"{prefix}_{stem}_{batch_id}.xlsx"
        if p.resolve() in inputs:
            raise ValueError("Output path would overwrite an input file")
        return p

    header, widths = read_template_header(opts.template, cfg.profile)
    if not opts.dry_run:
        p = out_path("FUEL_MASTERLIST_IMPORT")
        for rec in write_masterlist(p, header, widths, exported, cfg.profile):
            rec.add("FORMULA_TEXT_ESCAPED", cfg.policy.sev("FORMULA_TEXT_ESCAPED"), "", "Text starting with '=' written as literal text")
        result.outputs["import"] = p
    if opts.leads_log:
        p = out_path("FUEL_LEADS_LOG")
        write_leads_log(p, [leads_log_row(r.values, r.seller_slot, cfg.profile) for r in exported], cfg.profile)
        result.outputs["leads_log"] = p
    p = out_path("FUEL_IMPORT_AUDIT")
    audit.write_audit(p, result, opts)
    result.outputs["audit"] = p

    if not opts.dry_run:
        try:
            ledger.record_export({
                "batch_id": batch_id, "source_file_name": Path(opts.source).name, "source_file_sha256": source_sha,
                "source_sheet": sd.sheet_name, "mapping_profile_id": cfg.profile.raw.get("profile_id", ""),
                "responses_read": counts["responses"], "records_built": counts["records_built"],
                "records_exported": len(exported), "records_review": counts["records_review"],
                "records_skipped": counts["records_already_imported"], "output_path": str(result.outputs["import"]),
            }, [(r.src_key, r.source_row, r.dest_id, r.content_hash) for r in exported])
        except Exception:
            result.outputs["import"].unlink(missing_ok=True)  # never leave an un-ledgered import file
            raise
    log.info("batch %s: %s", batch_id, {k: v for k, v in counts.items() if k.startswith("records_")})
    return result


def run_from_paths(source: str | Path, template: str | Path, out_dir: str | Path, **kw) -> RunResult:
    return run(RunOptions(Path(source), Path(template), Path(out_dir), **kw))
