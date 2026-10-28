"""Load and validate the YAML/CSV configuration (design F.2, N.1, J.1)."""

from __future__ import annotations

import csv
import hashlib
import re
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from .model import Severity

VERIFIED = "VERIFIED"
OWNER_APPROVED_UNVERIFIED = "OWNER_APPROVED_UNVERIFIED"
UNRESOLVED = "UNRESOLVED"
_STATUSES = {VERIFIED, OWNER_APPROVED_UNVERIFIED, UNRESOLVED}
_VERIFIED_EVIDENCE = {"masterlist_export", "owner_confirmation"}


class ConfigError(ValueError):
    """Raised when configuration is missing or violates the approved rules."""


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _load_yaml(path: Path) -> dict:
    if not path.exists():
        raise ConfigError(f"Configuration file not found: {path}")
    with open(path, encoding="utf-8") as fh:
        data = yaml.safe_load(fh) or {}
    if not isinstance(data, dict):
        raise ConfigError(f"{path.name}: expected a mapping at the top level")
    return data


def normalise_label(label: object) -> str:
    return re.sub(r"\s+", " ", str(label or "")).strip().upper()


@dataclass
class TeamEntry:
    label_prefix: str
    slot: int | None
    status: str
    evidence_type: str = "none"
    evidence_ref: str = ""
    approved_by: str = ""
    approved_on: str = ""
    destination_label: str = ""
    fixture_slot: int | None = None

    @property
    def exportable(self) -> bool:
        return self.status in (VERIFIED, OWNER_APPROVED_UNVERIFIED)

    def describe(self) -> str:
        if self.status == VERIFIED:
            return f"Verified ({self.evidence_type}: {self.evidence_ref})"
        if self.status == OWNER_APPROVED_UNVERIFIED:
            return f"Owner-approved, not verified (by {self.approved_by} on {self.approved_on})"
        return "Unresolved"


@dataclass
class Teams:
    entries: list[TeamEntry]
    sha256: str = ""

    def lookup(self, label: object) -> TeamEntry | None:
        norm = normalise_label(label)
        hits = [e for e in self.entries if norm == normalise_label(e.label_prefix) or norm.startswith(normalise_label(e.label_prefix))]
        if len(hits) > 1:
            raise ConfigError(f"SALES LEADER label matches {len(hits)} teams.yaml entries; make the prefixes unique")
        return hits[0] if hits else None


def load_teams(path: Path) -> Teams:
    data = _load_yaml(path)
    entries = []
    for i, raw in enumerate(data.get("teams") or [], start=1):
        where = f"teams.yaml entry {i}"
        prefix = str(raw.get("label_prefix") or "").strip()
        status = str(raw.get("status") or "").strip()
        slot = raw.get("slot")
        if not prefix:
            raise ConfigError(f"{where}: label_prefix is required")
        if status not in _STATUSES:
            raise ConfigError(f"{where}: status must be one of {sorted(_STATUSES)}")
        if slot is not None and slot not in (1, 2, 3):
            raise ConfigError(f"{where}: slot must be 1, 2, 3 or null")
        e = TeamEntry(
            label_prefix=prefix, slot=slot, status=status,
            evidence_type=str(raw.get("evidence_type") or "none"),
            evidence_ref=str(raw.get("evidence_ref") or ""),
            approved_by=str(raw.get("approved_by") or ""),
            approved_on=str(raw.get("approved_on") or ""),
            destination_label=str(raw.get("destination_label") or ""),
            fixture_slot=raw.get("fixture_slot"),
        )
        if e.status == UNRESOLVED and e.slot is not None:
            raise ConfigError(f"{where}: an UNRESOLVED entry must have slot: null")
        if e.exportable and e.slot is None:
            raise ConfigError(f"{where}: {e.status} requires a slot (1, 2 or 3)")
        if e.status == VERIFIED and (e.evidence_type not in _VERIFIED_EVIDENCE or not e.evidence_ref):
            raise ConfigError(f"{where}: VERIFIED requires evidence_type in {sorted(_VERIFIED_EVIDENCE)} and an evidence_ref")
        if e.status == OWNER_APPROVED_UNVERIFIED and (not e.approved_by or not e.approved_on):
            raise ConfigError(f"{where}: OWNER_APPROVED_UNVERIFIED requires approved_by and approved_on")
        entries.append(e)
    prefixes = [normalise_label(e.label_prefix) for e in entries]
    if len(set(prefixes)) != len(prefixes):
        raise ConfigError("teams.yaml: duplicate label_prefix entries")
    return Teams(entries, file_sha256(path))


@dataclass
class Policy:
    severity: dict[str, Severity]
    action: dict[str, str]
    sha256: str = ""

    def sev(self, code: str) -> Severity:
        if code not in self.severity:
            raise ConfigError(f"validation_policy.yaml has no entry for issue code {code}")
        return self.severity[code]


def load_policy(path: Path) -> Policy:
    data = _load_yaml(path).get("codes") or {}
    sev, act = {}, {}
    for code, spec in data.items():
        try:
            sev[code] = Severity(str(spec.get("severity")))
        except ValueError as exc:
            raise ConfigError(f"validation_policy.yaml: {code} has an invalid severity") from exc
        act[code] = str(spec.get("action") or "")
    return Policy(sev, act, file_sha256(path))


@dataclass
class Vocab:
    projects: list[str]
    awareness: list[str]
    lead_sources: list[str]
    lead_source_v1: dict
    null_tokens: set[str]
    sha256: str = ""


def load_vocab(path: Path) -> Vocab:
    d = _load_yaml(path)
    return Vocab(
        projects=list(d.get("projects") or []),
        awareness=list(d.get("awareness") or []),
        lead_sources=list(d.get("lead_sources") or []),
        lead_source_v1=dict(d.get("lead_source_v1") or {}),
        null_tokens=set(d.get("null_tokens") or []),
        sha256=file_sha256(path),
    )


@dataclass
class Profile:
    raw: dict
    sha256: str = ""

    @property
    def source(self) -> dict:
        return self.raw["source"]

    @property
    def dest(self) -> dict:
        return self.raw["destination"]

    @property
    def dest_columns(self) -> list[str]:
        return list(self.dest["columns"])

    @property
    def lead_source_detail(self) -> dict[str, str]:
        return dict(self.raw.get("lead_source_detail") or {})

    @property
    def leads_log(self) -> dict:
        return self.raw["leads_log"]


def load_profile(path: Path) -> Profile:
    d = _load_yaml(path)
    for key in ("source", "destination", "leads_log"):
        if key not in d:
            raise ConfigError(f"{path.name}: missing section '{key}'")
    if len(d["destination"].get("columns") or []) != 35:
        raise ConfigError(f"{path.name}: destination.columns must list exactly 35 fields")
    return Profile(d, file_sha256(path))


@dataclass
class SellerDirectory:
    """Approved seller name -> e-mail. Unapproved rows are ignored (J.1)."""

    emails: dict[str, str] = field(default_factory=dict)
    source: str = "none"
    sha256: str = ""

    def get(self, seller_name: str) -> str | None:
        return self.emails.get(normalise_label(seller_name))


def load_sellers(path: Path | None) -> SellerDirectory:
    if path is None or not path.exists():
        return SellerDirectory()
    rows: list[dict] = []
    if path.suffix.lower() == ".csv":
        with open(path, encoding="utf-8-sig", newline="") as fh:
            for r in csv.DictReader(fh):
                rows.append({
                    "seller_name": r.get("seller_name", ""),
                    "email": (r.get("reviewer_email") or "").strip() or (r.get("proposed_email") or "").strip(),
                    "status": r.get("status", ""),
                })
    else:
        rows = list(_load_yaml(path).get("sellers") or [])
    emails = {}
    for r in rows:
        if str(r.get("status", "")).strip().upper() == "APPROVED" and r.get("email"):
            emails[normalise_label(r.get("seller_name"))] = str(r["email"]).strip()
    return SellerDirectory(emails, path.name, file_sha256(path))


@dataclass
class Config:
    profile: Profile
    vocab: Vocab
    policy: Policy
    teams: Teams
    sellers: SellerDirectory


def load_config(config_dir: Path, sellers_path: Path | None = None, teams_path: Path | None = None) -> Config:
    config_dir = Path(config_dir)
    if sellers_path is None:
        for candidate in ("sellers.csv", "sellers.yaml"):
            if (config_dir / candidate).exists():
                sellers_path = config_dir / candidate
                break
    return Config(
        profile=load_profile(config_dir / "mapping_profile.yaml"),
        vocab=load_vocab(config_dir / "vocab.yaml"),
        policy=load_policy(config_dir / "validation_policy.yaml"),
        teams=load_teams(teams_path or config_dir / "teams.yaml"),
        sellers=load_sellers(sellers_path),
    )
