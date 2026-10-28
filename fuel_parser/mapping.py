"""Resolve source columns to fields and build the mapping preview (design C, D.1, F.4)."""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.utils import get_column_letter

from .config import Profile
from .headers import match_key, suggest


class MappingError(ValueError):
    """Raised when the source structure cannot be mapped safely."""


@dataclass
class Block:
    ordinal: int  # 1-based order of the block in the sheet
    columns: dict[str, int]  # block field id -> column index


@dataclass
class PreviewItem:
    column: int
    header: str
    kind: str  # MATCHED | SUGGESTED | UNMAPPED | IGNORED_EMPTY
    target: str
    note: str = ""

    @property
    def letter(self) -> str:
        return get_column_letter(self.column + 1)


@dataclass
class ColumnMap:
    response: dict[str, int] = field(default_factory=dict)
    details: dict[str, list[int]] = field(default_factory=dict)
    others: dict[str, int] = field(default_factory=dict)
    blocks: list[Block] = field(default_factory=list)
    preview: list[PreviewItem] = field(default_factory=list)
    signature: str = ""

    @property
    def needs_confirmation(self) -> list[PreviewItem]:
        return [p for p in self.preview if p.kind in ("SUGGESTED", "UNMAPPED")]


def header_signature(headers: list[object]) -> str:
    joined = "\x1f".join(match_key(h) for h in headers)
    return hashlib.sha256(joined.encode("utf-8")).hexdigest()[:16]


def _detect_blocks(keys: list[str], block_cfg: dict) -> tuple[list[Block], set[int]]:
    start = block_cfg["start_key"]
    field_by_key = {v: k for k, v in block_cfg["fields"].items()}
    blocks, used = [], set()
    i = 0
    while i < len(keys):
        if keys[i] != start:
            i += 1
            continue
        cols: dict[str, int] = {}
        j = i
        while j < len(keys) and keys[j] in field_by_key and (j == i or keys[j] != start):
            fid = field_by_key[keys[j]]
            if fid in cols:
                break
            cols[fid] = j
            j += 1
        missing = sorted(set(field_by_key.values()) - set(cols))
        if missing:
            raise MappingError(
                f"BLOCK_STRUCTURE_ERROR: client block starting at column {get_column_letter(i + 1)} "
                f"is missing {', '.join(missing)}"
            )
        blocks.append(Block(len(blocks) + 1, cols))
        used.update(cols.values())
        i = j
    if not blocks:
        raise MappingError("No client block found (no 'Seller Name' header)")
    return blocks, used


def resolve(headers: list[object], profile: Profile, nonempty: dict[int, int],
            decisions: dict[str, str] | None = None) -> ColumnMap:
    """Map columns to fields.

    `nonempty` gives the count of non-blank values per column, so empty unknown columns
    can be listed without blocking. `decisions` holds confirmed choices from a saved
    mapping profile: header-key -> field id or "ignore".
    """
    src = profile.source
    keys = [match_key(h) for h in headers]
    cmap = ColumnMap(signature=header_signature(headers))
    decisions = decisions or {}

    blocks, used = _detect_blocks(keys, src["block"])
    cmap.blocks = blocks
    for b in blocks:
        for fid, col in b.columns.items():
            cmap.preview.append(PreviewItem(col, str(headers[col]), "MATCHED", f"block {b.ordinal}: {fid}"))

    for fid, aliases in src["response_fields"].items():
        hits = [i for i, k in enumerate(keys) if k in aliases and i not in used]
        if not hits:
            raise MappingError(f"Required source column for '{fid}' not found (aliases: {aliases})")
        if len(hits) > 1:
            raise MappingError(f"Source column for '{fid}' is ambiguous: {len(hits)} columns match")
        cmap.response[fid] = hits[0]
        used.add(hits[0])
        cmap.preview.append(PreviewItem(hits[0], str(headers[hits[0]]), "MATCHED", fid))

    # "Others" columns: the column immediately after the anchor question.
    for spec in src.get("others") or []:
        anchors = [i for i, k in enumerate(keys) if k == spec["after"]]
        if not anchors:
            continue
        nxt = anchors[0] + 1
        if nxt < len(keys) and keys[nxt] == spec["key"] and nxt not in used:
            cmap.others[spec["id"]] = nxt
            used.add(nxt)
            cmap.preview.append(PreviewItem(nxt, str(headers[nxt]), "MATCHED", f"{spec['id']} -> REMARKS [{spec['label']} – Others]"))

    for fid, aliases in src["details"].items():
        for i, k in enumerate(keys):
            if k in aliases and i not in used:
                cmap.details.setdefault(fid, []).append(i)
                used.add(i)
                cmap.preview.append(PreviewItem(i, str(headers[i]), "MATCHED", fid))

    all_aliases = sorted({a for v in src["details"].values() for a in v}
                         | {a for v in src["response_fields"].values() for a in v})
    for i, k in enumerate(keys):
        if i in used:
            continue
        h = "" if headers[i] is None else str(headers[i])
        decided = decisions.get(f"{i}:{k}")
        if decided == "ignore":
            cmap.preview.append(PreviewItem(i, h, "IGNORED_EMPTY" if not nonempty.get(i) else "MATCHED", "ignore", "confirmed: ignore"))
            continue
        if decided:
            if decided not in src["details"]:
                raise MappingError(f"Saved decision maps column {get_column_letter(i + 1)} to unknown field '{decided}'")
            cmap.details.setdefault(decided, []).append(i)
            cmap.preview.append(PreviewItem(i, h, "MATCHED", decided, "confirmed by user"))
            continue
        if not nonempty.get(i):
            cmap.preview.append(PreviewItem(i, h, "IGNORED_EMPTY", "", "no values"))
            continue
        guess = suggest(k, all_aliases)
        if guess:
            target = next(f for f, a in src["details"].items() if guess in a) if any(guess in a for a in src["details"].values()) else ""
            if target:
                cmap.preview.append(PreviewItem(i, h, "SUGGESTED", target, f"similar to '{guess}'"))
                continue
        cmap.preview.append(PreviewItem(i, h, "UNMAPPED", "", f"{nonempty[i]} values"))
    cmap.preview.sort(key=lambda p: p.column)
    return cmap


def decisions_from_preview(cmap: ColumnMap, headers: list[object], accept_suggestions: bool,
                           ignore_unmapped: bool) -> dict[str, str]:
    """Build explicit decisions for the items that need confirmation."""
    out = {}
    for p in cmap.needs_confirmation:
        key = f"{p.column}:{match_key(headers[p.column])}"
        if p.kind == "SUGGESTED" and accept_suggestions:
            out[key] = p.target
        elif p.kind == "UNMAPPED" and ignore_unmapped:
            out[key] = "ignore"
        elif p.kind == "SUGGESTED" and ignore_unmapped and not accept_suggestions:
            out[key] = "ignore"
    return out


def load_decisions(profiles_dir: Path, signature: str) -> dict[str, str]:
    p = Path(profiles_dir) / f"{signature}.json"
    if not p.exists():
        return {}
    return dict(json.loads(p.read_text(encoding="utf-8")).get("decisions") or {})


def save_decisions(profiles_dir: Path, signature: str, decisions: dict[str, str]) -> Path:
    d = Path(profiles_dir)
    d.mkdir(parents=True, exist_ok=True)
    existing = load_decisions(d, signature)
    existing.update(decisions)
    p = d / f"{signature}.json"
    p.write_text(json.dumps({"signature": signature, "decisions": existing}, indent=2), encoding="utf-8")
    return p
