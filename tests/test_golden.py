"""Golden runs on the real v3.3 export (T52, T53).

Skipped unless both environment variables point to local copies of the workbooks:
  FUEL_GOLDEN_SOURCE=/path/to/FUEL RESP (CSP-AD MARF) v3.3.xlsx
  FUEL_GOLDEN_TEMPLATE=/path/to/FUEL (MGM-AD ADMIN) (2).xlsx
Assertions use counts only; no workbook value is printed or stored in the repository.
"""

import hashlib
import os
from collections import Counter
from pathlib import Path

import openpyxl
import pytest

from fuel_parser.model import Severity
from fuel_parser.pipeline import RunOptions, run

REPO = Path(__file__).resolve().parent.parent
SRC = os.environ.get("FUEL_GOLDEN_SOURCE")
TPL = os.environ.get("FUEL_GOLDEN_TEMPLATE")
pytestmark = pytest.mark.skipif(not (SRC and TPL), reason="golden workbooks not provided")

SIX = ["TEAM LESLIE (AAD:", "TEAM FERDINAND (AAD:", "TEAM BERNADETTE (AAD:", "AM: JESSA", "TEAM CHRISTINE (AAD:", "TEAM RALPH (AAD:"]


def _run(tmp_path, **kw):
    return run(RunOptions(Path(SRC), Path(TPL), tmp_path / "out", config_dir=REPO / "config",
                          ledger_path=tmp_path / "ledger.sqlite", profiles_dir=tmp_path / "profiles", **kw))


def test_T52_shipped_config_exports_nothing_and_accounts_for_everything(tmp_path):
    before = hashlib.sha256(Path(SRC).read_bytes()).hexdigest()
    r = _run(tmp_path)
    c = r.counts
    assert (c["responses"], c["records_built"], c["records_exported"], c["records_review"]) == (3064, 3064, 0, 3064)
    assert c["responses_multi_block"] == 0 and c["responses_no_client"] == 0
    assert all(ok for _, ok, _ in r.checks)
    assert sum("SLOT_UNRESOLVED" in rec.codes() for rec in r.records) == 3064
    assert len(r.cmap.blocks) == 13 and not r.cmap.needs_confirmation
    assert hashlib.sha256(Path(SRC).read_bytes()).hexdigest() == before
    wb = openpyxl.load_workbook(r.outputs["import"], read_only=True)
    assert wb["MASTERLIST"].max_row == 1  # header only
    rowlog = openpyxl.load_workbook(r.outputs["audit"], read_only=True)["RowLog"]
    assert sum(1 for _ in rowlog.iter_rows(min_row=2)) == 3064


def test_team_sheet_cross_check_has_no_mismatches(tmp_path):
    r = _run(tmp_path, dry_run=True)
    ws = openpyxl.load_workbook(r.outputs["audit"], read_only=True)["TeamSheetCheck"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert rows[0][1] >= 3040  # team-sheet rows matched to a record
    assert all((row[2] or 0) == 0 for row in rows[2:])  # no field mismatches


def test_T53_owner_approved_test_config(tmp_path):
    teams = tmp_path / "teams.yaml"
    teams.write_text("teams:\n" + "".join(
        f'  - {{label_prefix: "{p}", slot: 1, status: OWNER_APPROVED_UNVERIFIED, evidence_type: business_rule, '
        f'approved_by: "TEST ONLY", approved_on: "2026-10-28"}}\n' for p in SIX), encoding="utf-8")
    r = _run(tmp_path, teams_path=teams, last_id=0)
    c = r.counts
    errors = Counter(i.code for rec in r.records if rec.status.value == "REVIEW" for i in rec.issues if i.severity == Severity.ERROR)
    assert errors == Counter({"UNMAPPED_VALUE": 23, "CLIENT_NAME_MISSING": 21})
    assert (c["records_exported"], c["records_review"], c["records_duplicate_held"]) == (3020, 44, 0)
    assert all("SLOT_OWNER_APPROVED_UNVERIFIED" in rec.codes() for rec in r.exported)
    # Re-running the same export must not re-export anything (safe re-runs).
    again = _run(tmp_path, teams_path=teams, last_id=3020)
    assert again.counts["records_exported"] == 0 and again.counts["records_already_imported"] == 3020
