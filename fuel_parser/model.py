"""Core data model shared by all pipeline stages."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum


class Severity(str, Enum):
    ERROR = "ERROR"
    HELD = "HELD"
    WARNING = "WARNING"
    INFO = "INFO"


class RecordStatus(str, Enum):
    EXPORTED = "EXPORTED"
    REVIEW = "REVIEW"
    DUPLICATE_HELD = "DUPLICATE_HELD"
    ALREADY_IMPORTED = "ALREADY_IMPORTED"
    EXCLUDED_BY_FILTER = "EXCLUDED_BY_FILTER"


@dataclass
class Issue:
    code: str
    severity: Severity
    source_row: int
    field: str
    message: str
    record_key: str = ""
    source_file: str = ""


@dataclass
class SourceResponse:
    """One data row of the response sheet, values keyed by column index."""

    row: int
    values: dict[int, object]


@dataclass
class Record:
    """One destination record built from one filled client block."""

    source_row: int
    block_index: int | None
    source_file: str = ""
    values: dict[str, object] = field(default_factory=dict)
    meta: dict[str, object] = field(default_factory=dict)
    issues: list[Issue] = field(default_factory=list)
    src_key: str = ""
    content_hash: str = ""
    client_key: str = ""
    duplicate_group: str = ""
    lead_source_original: str = ""
    sales_leader: str = ""
    seller_slot: int | None = None
    slot_status: str = ""
    timestamp: datetime | None = None
    status: RecordStatus | None = None
    dest_id: int | None = None

    def add(self, code: str, severity: Severity, fieldname: str, message: str) -> None:
        self.issues.append(Issue(code, severity, self.source_row, fieldname, message))

    def has(self, severity: Severity) -> bool:
        return any(i.severity == severity for i in self.issues)

    def codes(self) -> list[str]:
        return [i.code for i in self.issues]
