"""Local Streamlit UI (design F.4). Thin layer over fuel_parser.pipeline; all data stays on this computer.

Start with:  streamlit run app/streamlit_app.py
"""

from __future__ import annotations

import tempfile
from collections import Counter
from pathlib import Path

import streamlit as st

from fuel_parser.audit import mapping_status_rows
from fuel_parser.config import ConfigError
from fuel_parser.history import Ledger
from fuel_parser.mapping import MappingError
from fuel_parser.model import RecordStatus, Severity
from fuel_parser.pipeline import NeedsConfirmation, ReconciliationError, RunOptions, prepare, run
from fuel_parser.reader import ReadError
from fuel_parser.writer import TemplateError

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "output"
LEDGER = OUT_DIR / "import_history.sqlite"
ERRORS = (ConfigError, MappingError, ReadError, TemplateError, ReconciliationError, ValueError)

st.set_page_config(page_title="FUEL Response Parser", layout="wide")
st.title("FUEL Response Parser")
st.caption("Works only on files you select. Inputs are never modified; outputs are new files in the output folder.")


def _save_upload(upload) -> Path | None:
    if upload is None:
        return None
    d = Path(st.session_state.setdefault("tmpdir", tempfile.mkdtemp(prefix="fuel_")))
    p = d / upload.name
    p.write_bytes(upload.getvalue())
    return p


st.header("① Files")
c1, c2 = st.columns(2)
src = _save_upload(c1.file_uploader("Source export (.xlsx or .csv)", type=["xlsx", "csv"]))
tpl = _save_upload(c2.file_uploader("Destination template: FUEL (MGM-AD ADMIN).xlsx", type=["xlsx"]))
if not (src and tpl):
    st.info("Select both files to continue.")
    st.stop()

st.header("② Sheet and mapping")
sheet = st.text_input("Response sheet (leave blank to detect)", value="") or None
ts_fmt = st.text_input("CSV timestamp format (only for .csv with ambiguous dates)", value="") or None
accept = st.checkbox("Confirm suggested header matches")
ignore = st.checkbox("Confirm that listed unmapped columns are ignored")
base = dict(source=src, template=tpl, out_dir=OUT_DIR, ledger_path=LEDGER, profiles_dir=ROOT / "profiles",
            config_dir=ROOT / "config", sheet=sheet, timestamp_formats=[ts_fmt] if ts_fmt else None,
            accept_suggestions=accept, ignore_unmapped=ignore)
try:
    _, sd, cmap = prepare(RunOptions(**base))
except NeedsConfirmation as e:
    st.warning(str(e))
    if e.cmap:
        st.dataframe([{"Column": p.letter, "Header": p.header, "Result": p.kind, "Target": p.target, "Note": p.note}
                      for p in e.cmap.needs_confirmation], use_container_width=True)
    st.stop()
except ERRORS as e:
    st.error(str(e))
    st.stop()
st.success(f"Sheet '{sd.sheet_name}': {len(sd.rows)} responses, {len(cmap.blocks)} client blocks, mapping complete.")
with st.expander("Mapping preview"):
    st.dataframe([{"Column": p.letter, "Header": p.header, "Result": p.kind, "Target": p.target, "Note": p.note}
                  for p in cmap.preview], use_container_width=True)

st.header("③ Validate (dry run)")
teams = st.text_input("Export only these SALES LEADER prefixes (comma-separated, blank = all)", value="")
team_filter = [t.strip() for t in teams.split(",") if t.strip()] or None
if st.button("Validate"):
    try:
        st.session_state["dry"] = run(RunOptions(**base, team_filter=team_filter, dry_run=True))
    except ERRORS as e:
        st.error(str(e))
r = st.session_state.get("dry")
if r:
    c = r.counts
    cols = st.columns(5)
    for col, (label, key) in zip(cols, [("Responses", "responses"), ("Ready to export", "records_exported"),
                                        ("Held for review", "records_review"), ("Duplicates held", "records_duplicate_held"),
                                        ("Already imported", "records_already_imported")]):
        col.metric(label, c[key])
    if r.previous_batches:
        st.warning("This exact source file was exported before: " + ", ".join(b.batch_id for b in r.previous_batches))
    st.write("Reconciliation: " + " · ".join(("✅ " if ok else "❌ ") + name for name, ok, _ in r.checks))
    st.subheader("Mapping status")
    st.dataframe([dict(zip(["SALES LEADER", "Status", "Slot", "Evidence / approval", "Exported", "Held"], row))
                  for row in mapping_status_rows(r)], use_container_width=True)
    st.subheader("Issues")
    codes = Counter((i.code, i.severity.value) for rec in r.records for i in rec.issues)
    st.dataframe([{"Code": k, "Severity": s, "Count": n} for (k, s), n in sorted(codes.items())], use_container_width=True)
    st.subheader("Review (first 200)")
    review = [rec for rec in r.records if rec.status != RecordStatus.EXPORTED
              or any(i.severity == Severity.WARNING for i in rec.issues)][:200]
    st.dataframe([{"Status": rec.status.value, "Row": rec.source_row,
                   "Reasons": ", ".join(dict.fromkeys(i.code for i in rec.issues if i.severity != Severity.INFO)),
                   "SALES LEADER": rec.sales_leader, "Lead source": rec.lead_source_original} for rec in review],
                 use_container_width=True)
    if r.outputs.get("audit"):
        st.download_button("Download dry-run audit", r.outputs["audit"].read_bytes(), file_name=r.outputs["audit"].name)

    st.header("④ Export")
    ready = c["records_exported"]
    last_id = st.number_input("Last Id in the live MASTERLIST", min_value=0, step=1, value=0)
    leads_log = st.checkbox("Also write the optional 17-column Leads Log")
    if ready == 0:
        st.info("No records are ready to export. You can still export to get the full audit and review file.")
    if st.button(f"Export ({ready} records)"):
        try:
            res = run(RunOptions(**base, team_filter=team_filter, last_id=int(last_id), leads_log=leads_log))
            st.success(f"Batch {res.batch_id} written to {OUT_DIR}. Append the import file to the live MASTERLIST yourself.")
            for kind, path in res.outputs.items():
                st.download_button(f"Download {kind}", path.read_bytes(), file_name=path.name, key=kind)
        except ERRORS as e:
            st.error(str(e))

with st.sidebar:
    st.subheader("Import history")
    for b in Ledger(LEDGER).list_batches()[-10:]:
        st.write(f"{b.batch_id} · {b.status} · {b.records_exported} exported")
    vb = st.text_input("Void batch id")
    if st.button("Void batch") and vb:
        try:
            st.write(f"Released {Ledger(LEDGER).void_batch(vb)} record key(s).")
        except KeyError as e:
            st.error(str(e))
