# FUEL Response Parser

Converts FUEL Google Forms response exports (`.xlsx` or `.csv`) into the 35-column `MASTERLIST` import format of
`FUEL (MGM-AD ADMIN).xlsx`, with a full audit report. It works **only on local files**: no Google or Microsoft account,
API, credentials or network access is needed. Input files are never modified, and the live MASTERLIST is never written
to. You append the import file yourself.

Design and decision register: [`docs/DESIGN.md`](docs/DESIGN.md) (v0.7).

## Install (once)

Requires Python 3.11 or newer.

```bash
python -m venv .venv
.venv/bin/pip install -e ".[ui,dev]"
```

On Windows, use `.venv\Scripts\pip` and `.venv\Scripts\fuel-parser` instead of `.venv/bin/...`.

## First run

1. Put your exports somewhere outside the repository folder, or in `output/` (git-ignored).
2. Check the mapping:
   ```bash
   .venv/bin/fuel-parser preview --source "FUEL RESP (CSP-AD MARF) v3.3.xlsx" --template "FUEL (MGM-AD ADMIN) (2).xlsx"
   ```
3. Validate without exporting (writes only an audit file to `output/`):
   ```bash
   .venv/bin/fuel-parser run --source "FUEL RESP (CSP-AD MARF) v3.3.xlsx" --template "FUEL (MGM-AD ADMIN) (2).xlsx" --dry-run
   ```
4. Open `output/FUEL_IMPORT_AUDIT_….xlsx`: **Summary** (reconciliation), **MappingStatus**, **Review** (each held
   record with its reason and the action to take), **RowLog** (every source row), **TeamSheetCheck**.

With the shipped `config/teams.yaml`, every SALES LEADER is `UNRESOLVED`, so **0 records are exported** and all
responses are listed in Review with `SLOT_UNRESOLVED`. This is the intended starting state.

Or use the local UI: `.venv/bin/streamlit run app/streamlit_app.py` (opens on `localhost` only).

## Enabling a team (no code change)

Edit `config/teams.yaml` and change one entry:

```yaml
# Verified by evidence (a MASTERLIST export or written confirmation from the form owner):
- {label_prefix: "TEAM LESLIE (AAD:", slot: 1, status: VERIFIED, evidence_type: masterlist_export,
   evidence_ref: "MASTERLIST export 2026-11-01, sha256 <hash>"}
# Or your explicit operating assumption (exported with a warning, never reported as verified):
- {label_prefix: "TEAM LESLIE (AAD:", slot: 1, status: OWNER_APPROVED_UNVERIFIED, evidence_type: business_rule,
   approved_by: "<your name>", approved_on: "2026-11-01"}
```

`fuel-parser preflight-masterlist --masterlist <export.xlsx>` counts SALES LEADER × seller slot in a MASTERLIST export
to support a `VERIFIED` entry. It never changes the configuration.

## Exporting

```bash
.venv/bin/fuel-parser run --source <export.xlsx> --template <template.xlsx> --last-id <last Id in the live MASTERLIST>
```

- `--leads-log` also writes the optional 17-column Leads Log.
- `--team "TEAM X"` limits the export to some teams. The others are counted as excluded.
- Exported record keys are stored in `output/import_history.sqlite` (hashes and row numbers only), so re-running the
  same or a newer export never re-exports a record.
- If you do not append a batch to the live MASTERLIST: `fuel-parser void-batch <batch id>` (list with `fuel-parser batches`).

## Seller directory (optional)

```bash
.venv/bin/fuel-parser draft-directory --source <export.xlsx> --out config/sellers.csv
```

Review the CSV and set `status` to `APPROVED` or `REJECTED` per row. Fill `reviewer_email` to override an address.
Only `APPROVED` rows are used; otherwise the submitter e-mail is used and logged. `config/sellers.csv` contains
personal data and is git-ignored.

## Other options

| Situation | Option |
|---|---|
| Workbook has no `MARF RAW` sheet | `--sheet "<name>"` after checking the proposed sheet |
| CSV with ambiguous dates | `--timestamp-format "%m/%d/%Y %H:%M:%S"` |
| Renamed headers flagged as suggestions | check them, then `--accept-suggestions` (saved per header layout in `profiles/`) |
| Unknown non-empty columns | check them, then `--ignore-unmapped` (saved) |

## Tests

```bash
.venv/bin/pytest
```

Synthetic fixtures only. To also run the golden checks against the real export, point to local copies:

```bash
FUEL_GOLDEN_SOURCE="/path/FUEL RESP (CSP-AD MARF) v3.3.xlsx" FUEL_GOLDEN_TEMPLATE="/path/FUEL (MGM-AD ADMIN) (2).xlsx" .venv/bin/pytest
```

## Privacy

Real workbooks, outputs, the ledger, the seller directory and saved profiles are git-ignored. Never commit them. The
audit and review files contain client data: keep them on your computer.
