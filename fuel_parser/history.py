"""Local import ledger (design D.5). Stores hashes and row numbers only, never client data."""

from __future__ import annotations

import sqlite3
from contextlib import closing
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

_SCHEMA = """
CREATE TABLE IF NOT EXISTS import_batch (
    batch_id TEXT PRIMARY KEY,
    created_at TEXT NOT NULL,
    source_file_name TEXT NOT NULL,
    source_file_sha256 TEXT NOT NULL,
    source_sheet TEXT NOT NULL,
    mapping_profile_id TEXT NOT NULL,
    responses_read INTEGER NOT NULL,
    records_built INTEGER NOT NULL,
    records_exported INTEGER NOT NULL,
    records_review INTEGER NOT NULL,
    records_skipped INTEGER NOT NULL,
    status TEXT NOT NULL,
    output_path TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS imported_record (
    src_key TEXT PRIMARY KEY,
    batch_id TEXT NOT NULL REFERENCES import_batch(batch_id),
    source_row INTEGER NOT NULL,
    dest_id INTEGER NOT NULL,
    content_hash TEXT NOT NULL,
    exported_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_imported_content ON imported_record(content_hash);
"""


@dataclass
class BatchRow:
    batch_id: str
    created_at: str
    source_file_name: str
    records_exported: int
    status: str
    output_path: str


class Ledger:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with closing(self._connect()) as con:
            con.executescript(_SCHEMA)

    def _connect(self) -> sqlite3.Connection:
        return sqlite3.connect(self.path)

    def imported_keys(self) -> set[str]:
        with closing(self._connect()) as con:
            return {r[0] for r in con.execute("SELECT src_key FROM imported_record")}

    def imported_content_hashes(self) -> set[str]:
        with closing(self._connect()) as con:
            return {r[0] for r in con.execute("SELECT content_hash FROM imported_record")}

    def batches_for_file(self, sha256: str) -> list[BatchRow]:
        with closing(self._connect()) as con:
            rows = con.execute(
                "SELECT batch_id, created_at, source_file_name, records_exported, status, output_path "
                "FROM import_batch WHERE source_file_sha256 = ? AND status = 'exported' ORDER BY created_at", (sha256,)).fetchall()
        return [BatchRow(*r) for r in rows]

    def list_batches(self) -> list[BatchRow]:
        with closing(self._connect()) as con:
            rows = con.execute(
                "SELECT batch_id, created_at, source_file_name, records_exported, status, output_path "
                "FROM import_batch ORDER BY created_at").fetchall()
        return [BatchRow(*r) for r in rows]

    def record_export(self, batch: dict, records: list[tuple[str, int, int, str]]) -> None:
        """Write the batch and its exported keys in one transaction.

        `records` holds (src_key, source_row, dest_id, content_hash).
        """
        now = datetime.now().isoformat(timespec="seconds")
        with closing(self._connect()) as con, con:
            con.execute(
                "INSERT INTO import_batch VALUES (:batch_id, :created_at, :source_file_name, :source_file_sha256, "
                ":source_sheet, :mapping_profile_id, :responses_read, :records_built, :records_exported, "
                ":records_review, :records_skipped, 'exported', :output_path)", {**batch, "created_at": now})
            con.executemany(
                "INSERT INTO imported_record VALUES (?, ?, ?, ?, ?, ?)",
                [(k, batch["batch_id"], row, did, ch, now) for k, row, did, ch in records])

    def void_batch(self, batch_id: str) -> int:
        """Release a batch's keys (its output was never appended to the live MASTERLIST)."""
        with closing(self._connect()) as con, con:
            cur = con.execute("SELECT status FROM import_batch WHERE batch_id = ?", (batch_id,)).fetchone()
            if cur is None:
                raise KeyError(f"Unknown batch: {batch_id}")
            if cur[0] == "voided":
                return 0
            n = con.execute("DELETE FROM imported_record WHERE batch_id = ?", (batch_id,)).rowcount
            con.execute("UPDATE import_batch SET status = 'voided' WHERE batch_id = ?", (batch_id,))
            return n
