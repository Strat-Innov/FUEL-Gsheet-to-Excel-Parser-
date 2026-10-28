"""Command-line interface (design F.1). Works entirely on local files."""

from __future__ import annotations

import argparse
import csv
import logging
import sys
from collections import Counter, defaultdict
from pathlib import Path

from .audit import mapping_status_rows
from .config import ConfigError, load_config, normalise_label
from .history import Ledger
from .mapping import MappingError
from .pipeline import NeedsConfirmation, ReconciliationError, RunOptions, prepare, run
from .reader import ReadError, read_sheet
from .writer import TemplateError

DEFAULT_LEDGER = Path("output/import_history.sqlite")


def _common(p: argparse.ArgumentParser) -> None:
    p.add_argument("--source", required=True, type=Path, help="Google Forms response export (.xlsx or .csv)")
    p.add_argument("--template", required=True, type=Path, help="FUEL (MGM-AD ADMIN) workbook with the MASTERLIST sheet")
    p.add_argument("--config-dir", type=Path, default=Path("config"))
    p.add_argument("--profiles-dir", type=Path, default=Path("profiles"), help="Saved mapping confirmations")
    p.add_argument("--sellers", type=Path, help="Reviewed seller directory (.csv or .yaml); default config/sellers.csv|yaml")
    p.add_argument("--teams", type=Path, help="Alternative teams.yaml")
    p.add_argument("--sheet", help="Response sheet name (confirms a proposed sheet)")
    p.add_argument("--timestamp-format", action="append", dest="timestamp_formats",
                   help="Confirmed strptime format for text timestamps in CSV exports, e.g. '%%m/%%d/%%Y %%H:%%M:%%S'")
    p.add_argument("--accept-suggestions", action="store_true", help="Confirm suggested header matches (saved per header signature)")
    p.add_argument("--ignore-unmapped", action="store_true", help="Confirm that listed unmapped columns are ignored (saved)")


def _options(a: argparse.Namespace, **extra) -> RunOptions:
    return RunOptions(
        source=a.source, template=a.template, out_dir=getattr(a, "out_dir", Path("output")),
        config_dir=a.config_dir, ledger_path=getattr(a, "ledger", DEFAULT_LEDGER), profiles_dir=a.profiles_dir,
        sellers_path=a.sellers, teams_path=a.teams, sheet=a.sheet, timestamp_formats=a.timestamp_formats,
        accept_suggestions=a.accept_suggestions, ignore_unmapped=a.ignore_unmapped, **extra)


def _print_preview(cmap) -> None:
    kinds = Counter(p.kind for p in cmap.preview)
    print(f"Header signature {cmap.signature}: {len(cmap.blocks)} client blocks; "
          + ", ".join(f"{k} {v}" for k, v in sorted(kinds.items())))
    for p in cmap.preview:
        if p.kind != "MATCHED":
            print(f"  {p.letter:>3}  {p.kind:<13} {p.header[:60]!r:<64} {p.target} {p.note}")


def cmd_preview(a) -> int:
    try:
        _, sd, cmap = prepare(_options(a))
    except NeedsConfirmation as e:
        print(f"NEEDS CONFIRMATION: {e}")
        if e.cmap:
            _print_preview(e.cmap)
        print("Re-run with --sheet, --accept-suggestions and/or --ignore-unmapped after checking the items above.")
        return 2
    print(f"Response sheet '{sd.sheet_name}': {len(sd.rows)} responses, {sd.blank_rows} blank rows"
          + (f"; ignored: {', '.join(sd.ignored_sheets)}" if sd.ignored_sheets else ""))
    _print_preview(cmap)
    print("Mapping complete: no item needs confirmation.")
    return 0


def cmd_run(a) -> int:
    opts = _options(a, last_id=a.last_id, team_filter=a.team, leads_log=a.leads_log, dry_run=a.dry_run,
                    team_sheet_check=a.team_sheet_check)
    try:
        r = run(opts)
    except NeedsConfirmation as e:
        print(f"NEEDS CONFIRMATION: {e}")
        if e.cmap:
            _print_preview(e.cmap)
        print("Nothing was written. Check the items, then re-run with --sheet / --accept-suggestions / --ignore-unmapped.")
        return 2
    if r.previous_batches:
        print("WARNING: this exact source file was exported before: "
              + ", ".join(f"{b.batch_id} ({b.records_exported} exported)" for b in r.previous_batches))
    c = r.counts
    print(f"{'DRY RUN — ' if a.dry_run else ''}Batch {r.batch_id}")
    print(f"  Response sheet: '{r.sheet.sheet_name}'" + (f" ({len(r.sheet.ignored_sheets)} other sheet(s) ignored)" if r.sheet.ignored_sheets else ""))
    print(f"  Responses read: {c['responses']}  (blank rows {c['blank_rows']}, no client data {c['responses_no_client']}, "
          f"multi-block {c['responses_multi_block']})")
    print(f"  Records built: {c['records_built']} = exported {c['records_exported']} + review {c['records_review']} + "
          f"duplicates held {c['records_duplicate_held']} + already imported {c['records_already_imported']} + "
          f"excluded by filter {c['records_excluded_by_filter']}")
    for name, ok, _ in r.checks:
        print(f"  [{'PASS' if ok else 'FAIL'}] {name}")
    print("  Mapping status:")
    for leader, status, slot, desc, exp, held in mapping_status_rows(r):
        print(f"    {leader[:45]:<45} {status:<26} slot {slot or '-'}  exported {exp:>5}  held {held:>5}")
    codes = Counter(i.code for rec in r.records for i in rec.issues) + Counter(i.code for i in r.response_issues)
    print("  Issues: " + ", ".join(f"{k} {v}" for k, v in sorted(codes.items())))
    for kind, path in r.outputs.items():
        print(f"  {kind:<9} {path}")
    if not a.dry_run and c["records_exported"]:
        print("Next: check the Review sheet, then append the import file to the live MASTERLIST yourself. "
              f"If you do not use it, run: fuel-parser void-batch {r.batch_id}")
    return 0


def cmd_batches(a) -> int:
    for b in Ledger(a.ledger).list_batches():
        print(f"{b.batch_id}  {b.created_at}  {b.status:<8} exported {b.records_exported:>5}  {b.source_file_name}")
    return 0


def cmd_void(a) -> int:
    n = Ledger(a.ledger).void_batch(a.batch_id)
    print(f"Batch {a.batch_id} voided: {n} record key(s) released.")
    return 0


def cmd_preflight(a) -> int:
    """Count SALES LEADER x filled seller slot in a MASTERLIST export (evidence for M.4)."""
    cfg = load_config(a.config_dir)
    sd = read_sheet(a.masterlist, cfg.profile.dest["sheet"])
    from .headers import match_key
    keys = [match_key(h) for h in sd.headers]
    idx = {k: i for i, k in enumerate(keys)}
    pairs = {1: ("seller name", "seller's email address"), 2: ("seller name1", "seller's email address1"),
             3: ("seller name2", "seller's email address2")}
    counts: dict[str, Counter] = defaultdict(Counter)
    for _, row in sd.rows:
        leader = normalise_label(row[idx["sales leader"]])
        filled = [n for n, (nk, ek) in pairs.items() if row[idx[nk]] or row[idx[ek]]]
        counts[leader][tuple(filled) if filled else ("none",)] += 1
    print(f"MASTERLIST export: {len(sd.rows)} rows")
    for leader, c in sorted(counts.items()):
        slots = {s for key in c for s in key if s != "none"}
        total = sum(c.values())
        verdict = ("CONSISTENT" if len(slots) == 1 and total >= 3 and all(len(k) == 1 for k in c)
                   else "CONTRADICTION" if len(slots) > 1 else "INSUFFICIENT")
        print(f"  {leader[:45]:<45} rows {total:>5}  slots {dict(c)}  -> {verdict}")
    print("CONSISTENT leaders may be recorded in teams.yaml as VERIFIED with evidence_type: masterlist_export "
          "and evidence_ref: <file name + sha256>. Nothing was changed automatically.")
    return 0


def cmd_draft_directory(a) -> int:
    """Draft the seller directory for human review (J.1). Output contains personal data: keep it local."""
    cfg = load_config(a.config_dir)
    from .mapping import resolve
    from .pipeline import _nonempty_counts
    from .reader import choose_sheet
    src = cfg.profile.source
    sheet = a.sheet or choose_sheet(a.source, src["preferred_sheet"], src["signature_keys"]).name
    sd = read_sheet(a.source, sheet)
    cmap = resolve(sd.headers, cfg.profile, _nonempty_counts(sd))
    sub_col = cmap.response["submitter_email"]
    by: dict[str, Counter] = defaultdict(Counter)
    sellers_of: dict[str, set] = defaultdict(set)
    for _, row in sd.rows:
        for b in cmap.blocks:
            name = row[b.columns["seller_name"]]
            if name and str(name).strip():
                n, e = str(name).strip(), str(row[sub_col] or "").strip().lower()
                by[n][e] += 1
                sellers_of[e].add(n)
    with open(a.out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["seller_name", "proposed_email", "responses", "share_pct", "other_emails", "email_also_used_for",
                    "review_flags", "status", "reviewer_email", "reviewer_note"])
        for n, c in sorted(by.items()):
            (e, k), *rest = c.most_common()
            share = round(100 * k / sum(c.values()), 1)
            flags = [f for f, on in (("TIE", bool(rest) and rest[0][1] == k), ("LOW_SHARE", share < 80),
                                     ("EMAIL_USED_FOR_OTHER_SELLERS", len(sellers_of[e]) > 1)) if on]
            w.writerow([n, e, sum(c.values()), share, " | ".join(f"{x} ({v})" for x, v in rest),
                        " | ".join(sorted(sellers_of[e] - {n})), ",".join(flags) or "OK", "DRAFT - NOT APPROVED", "", ""])
    print(f"Draft directory written to {a.out} ({len(by)} sellers). Set status to APPROVED or REJECTED per row, "
          "then save it as config/sellers.csv (git-ignored).")
    return 0


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.WARNING, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(prog="fuel-parser", description="FUEL response export -> MASTERLIST import file")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("preview", help="Show the sheet and mapping preview")
    _common(p)
    p.set_defaults(func=cmd_preview)

    p = sub.add_parser("run", help="Validate, transform and write the import and audit files")
    _common(p)
    p.add_argument("--out-dir", type=Path, default=Path("output"))
    p.add_argument("--ledger", type=Path, default=DEFAULT_LEDGER)
    p.add_argument("--last-id", type=int, help="Last Id in the live MASTERLIST (required when records are exported)")
    p.add_argument("--team", action="append", help="Export only these SALES LEADER prefixes (others are counted as excluded)")
    p.add_argument("--leads-log", action="store_true", help="Also write the optional 17-column Leads Log")
    p.add_argument("--dry-run", action="store_true", help="Audit only: no import file, ledger unchanged")
    p.add_argument("--team-sheet-check", action="store_true",
                   help="Also compare against the workbook's team sheets in the audit (reference only; off by default)")
    p.set_defaults(func=cmd_run)

    p = sub.add_parser("batches", help="List import batches in the ledger")
    p.add_argument("--ledger", type=Path, default=DEFAULT_LEDGER)
    p.set_defaults(func=cmd_batches)

    p = sub.add_parser("void-batch", help="Release a batch whose output was never appended to the live MASTERLIST")
    p.add_argument("batch_id")
    p.add_argument("--ledger", type=Path, default=DEFAULT_LEDGER)
    p.set_defaults(func=cmd_void)

    p = sub.add_parser("preflight-masterlist", help="Count SALES LEADER x seller slot in a MASTERLIST export")
    p.add_argument("--masterlist", required=True, type=Path)
    p.add_argument("--config-dir", type=Path, default=Path("config"))
    p.set_defaults(func=cmd_preflight)

    p = sub.add_parser("draft-directory", help="Draft the seller directory CSV for review")
    p.add_argument("--source", required=True, type=Path)
    p.add_argument("--out", required=True, type=Path)
    p.add_argument("--sheet")
    p.add_argument("--config-dir", type=Path, default=Path("config"))
    p.set_defaults(func=cmd_draft_directory)

    a = ap.parse_args(argv)
    try:
        return a.func(a)
    except (ConfigError, MappingError, ReadError, TemplateError, ReconciliationError, ValueError, KeyError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
