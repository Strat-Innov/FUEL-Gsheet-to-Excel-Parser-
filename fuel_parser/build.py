"""Build destination records from one source response (design C, D.1–D.3, J.4, J.5, N.2)."""

from __future__ import annotations

import hashlib
import json
import re
from datetime import datetime

from .config import OWNER_APPROVED_UNVERIFIED, Config
from .mapping import ColumnMap
from .model import Issue, Record, SourceResponse
from .transform import (classify_mobile, clean_text, is_valid_email, join_options,
                        parse_timestamp, tokenize_options)

# Detail fields written to the destination column of the same name.
_DIRECT_DETAILS = ("booth", "referrer", "digital_official", "digital_seller", "own_source",
                   "roadshow", "saturation", "showroom", "sponsorship", "seller_fb_page")
_DEST_DETAILS = _DIRECT_DETAILS + ("repeat_buyer",)
_CLIENT_TEXT = ("surname", "first_name", "middle_name", "client_email", "residence", "remarks")


def _sha(*parts: object) -> str:
    return hashlib.sha256("|".join("" if p is None else str(p) for p in parts).encode("utf-8")).hexdigest()


class RecordBuilder:
    def __init__(self, cfg: Config, cmap: ColumnMap, timestamp_formats: list[str] | None = None,
                 now: datetime | None = None):
        self.cfg = cfg
        self.cmap = cmap
        self.formats = timestamp_formats
        self.now = now
        self.others_specs = {s["id"]: s for s in cfg.profile.source.get("others") or []}

    # -- helpers ---------------------------------------------------------------------
    def _flag(self, rec: Record, code: str, fieldname: str, message: str) -> None:
        rec.add(code, self.cfg.policy.sev(code), fieldname, message)

    def _text(self, rec: Record, raw: object, fieldname: str, null_tokens: bool = False) -> object:
        c = clean_text(raw, self.cfg.vocab.null_tokens if null_tokens else None)
        if c.trimmed:
            self._flag(rec, "WHITESPACE_TRIMMED", fieldname, "Leading/trailing whitespace removed")
        if c.null_token:
            self._flag(rec, "NULL_TOKEN_BLANKED", fieldname, "Placeholder value written as blank")
        return c.value

    def block_is_filled(self, resp: SourceResponse, block) -> bool:
        for col in block.columns.values():
            c = clean_text(resp.values.get(col), self.cfg.vocab.null_tokens)
            if c.value is not None:
                return True
        return False

    # -- main ------------------------------------------------------------------------
    def build(self, resp: SourceResponse) -> tuple[list[Record], list[Issue]]:
        """Return the records for a response, or response-level issues when it has none."""
        filled = [b for b in self.cmap.blocks if self.block_is_filled(resp, b)]
        if not filled:
            sev = self.cfg.policy.sev("NO_CLIENT_DATA")
            return [], [Issue("NO_CLIENT_DATA", sev, resp.row, "", "Response has no filled client block")]
        records = []
        for block in filled:
            rec = self._build_one(resp, block)
            if len(filled) > 1:
                self._flag(rec, "MULTI_BLOCK_RESPONSE", "", f"Response fills {len(filled)} client blocks; this is block {block.ordinal}")
            records.append(rec)
        return records, []

    def _build_one(self, resp: SourceResponse, block) -> Record:
        cfg, cm = self.cfg, self.cmap
        v = resp.values
        rec = Record(source_row=resp.row, block_index=block.ordinal)
        out = {fid: None for fid in cfg.profile.dest_columns}

        # Response-level fields.
        ts = parse_timestamp(v.get(cm.response["timestamp"]), self.formats, self.now)
        if ts.code:
            self._flag(rec, ts.code, "completion_time", ts.message)
        rec.timestamp = ts.value
        out["start_time"] = out["completion_time"] = ts.value  # H-4

        submitter = self._text(rec, v.get(cm.response["submitter_email"]), "email")
        out["email"] = submitter  # H-3; Name stays blank
        leader = self._text(rec, v.get(cm.response["sales_leader"]), "sales_leader")
        rec.sales_leader = leader or ""

        # Block fields.
        bv = {fid: v.get(col) for fid, col in block.columns.items()}
        for fid in _CLIENT_TEXT:
            out[fid] = self._text(rec, bv.get(fid), fid, null_tokens=True)
        rec.meta["remarks_original"] = out["remarks"]
        for fid in ("surname", "first_name"):
            if not out[fid]:
                self._flag(rec, "CLIENT_NAME_MISSING", fid, "Required client name is blank")
        if out["client_email"] and not is_valid_email(str(out["client_email"])):
            self._flag(rec, "INVALID_EMAIL", "client_email", "Client e-mail syntax is invalid")

        raw_mobile = bv.get("mobile")
        m = classify_mobile(raw_mobile if not isinstance(raw_mobile, str) else clean_text(raw_mobile, cfg.vocab.null_tokens).value)
        out["mobile"] = m.text
        if m.text is None:
            self._flag(rec, "MOBILE_MISSING", "mobile", "Mobile is blank")
        if m.was_numeric:
            self._flag(rec, "MOBILE_WAS_NUMERIC", "mobile", "Mobile was stored as a number")
        if m.reasons:
            self._flag(rec, "MOBILE_FORMAT_UNEXPECTED", "mobile", f"Unexpected mobile format: {', '.join(m.reasons)}")

        self._projects(rec, out, bv)
        self._seller(rec, out, bv, submitter, leader)
        self._lead_source(rec, out, v)
        self._awareness(rec, out, v)
        self._others(rec, out, v)

        rec.values = out
        self._keys(rec, submitter)
        return rec

    def _projects(self, rec: Record, out: dict, bv: dict) -> None:
        known = self.cfg.vocab.projects
        lists = {}
        for fid in ("project_first", "project_second"):
            raw = self._text(rec, bv.get(fid), fid)
            if raw is None:
                lists[fid] = []
                continue
            tokens, unknown = tokenize_options(str(raw), known)
            for u in unknown:
                self._flag(rec, "UNKNOWN_PROJECT", fid, f"Project not in the known list: {u}")
            out[fid] = join_options(tokens)
            lists[fid] = tokens
        if len(lists["project_first"]) > 1:
            self._flag(rec, "MULTI_FIRST_CHOICE", "project_first", f"{len(lists['project_first'])} first-choice projects kept")
        if set(lists["project_first"]) & set(lists["project_second"]):
            self._flag(rec, "SECOND_INCLUDES_FIRST", "project_second", "Second choice repeats a first-choice project")

    def _seller(self, rec: Record, out: dict, bv: dict, submitter: object, leader: object) -> None:
        cfg = self.cfg
        seller = self._text(rec, bv.get("seller_name"), "seller_name")
        if not seller:
            self._flag(rec, "SELLER_NAME_MISSING", "seller_name", "Seller name is blank")
        directory_email = cfg.sellers.get(str(seller)) if seller else None
        if directory_email:
            email = directory_email
            if submitter and str(submitter).lower() != directory_email.lower():
                self._flag(rec, "SUBMITTER_NOT_SELLER", "seller_email", "Submitting account differs from the approved directory e-mail")
        else:
            email = submitter
            if email:
                self._flag(rec, "SELLER_EMAIL_FROM_SUBMITTER", "seller_email", "Seller e-mail taken from the submitting account")
        if not email:
            self._flag(rec, "SELLER_EMAIL_MISSING", "seller_email", "No seller e-mail available")
        elif not is_valid_email(str(email)):
            self._flag(rec, "INVALID_EMAIL", "seller_email", "Seller e-mail syntax is invalid")
        if submitter and submitter != email and not is_valid_email(str(submitter)):
            self._flag(rec, "INVALID_EMAIL", "email", "Submitter e-mail syntax is invalid")
        rec.meta["seller_name"], rec.meta["seller_email"] = seller, email

        team = cfg.teams.lookup(leader) if leader else None
        out["sales_leader"] = (team.destination_label if team and team.destination_label else leader)
        if team is None or not team.exportable:
            rec.slot_status = team.status if team else "UNRESOLVED (not in teams.yaml)"
            self._flag(rec, "SLOT_UNRESOLVED", "sales_leader",
                       "SALES LEADER has no exportable seller-slot mapping" + ("" if team else " (not listed in teams.yaml)"))
            return  # N.2: nothing is written to any seller slot
        rec.seller_slot, rec.slot_status = team.slot, team.status
        if seller:  # the name/e-mail pair is written together, never an e-mail alone (AC5)
            out[f"seller_name_{team.slot}"] = seller
            out[f"seller_email_{team.slot}"] = email
        if team.status == OWNER_APPROVED_UNVERIFIED:
            self._flag(rec, "SLOT_OWNER_APPROVED_UNVERIFIED", "sales_leader", team.describe())

    def _detail_value(self, rec: Record, v: dict, fid: str) -> object:
        values = []
        for col in self.cmap.details.get(fid, []):
            val = self._text(rec, v.get(col), fid)
            if val is not None:
                values.append(val)
        if len(set(map(str, values))) > 1:
            self._flag(rec, "CONFLICTING_SOURCE_VALUES", fid, "More than one source column holds a value for this field")
        return values[0] if values else None

    def _lead_source(self, rec: Record, out: dict, v: dict) -> None:
        cfg = self.cfg
        raw = self._text(rec, v.get(self.cmap.response["lead_source"]), "lead_source")
        rec.lead_source_original = raw or ""
        details = {fid: self._detail_value(rec, v, fid) for fid in self.cfg.profile.source["details"]}
        rec.meta["details_raw"] = details

        for fid in _DIRECT_DETAILS:
            out[fid] = details.get(fid)
        if details.get("repeat_buyer"):
            tokens, _ = tokenize_options(str(details["repeat_buyer"]), cfg.vocab.projects)
            out["repeat_buyer"] = join_options(tokens)
        if details.get("corporate"):
            self._flag(rec, "UNMAPPED_COLUMN", "question", "Corporate detail has a value but `Question` stays blank (B3)")

        ls = None
        if not raw:
            self._flag(rec, "MISSING_LEAD_SOURCE", "lead_source", "Lead Source is blank")
        elif raw in cfg.vocab.lead_sources:
            ls = raw
        elif raw in cfg.vocab.lead_source_v1:
            rule = cfg.vocab.lead_source_v1[raw]
            if "to" in rule:
                ls = rule["to"]
            else:
                sub = details.get(rule.get("by_detail", ""))
                hit = (rule.get("when") or {}).get(sub) if sub else None
                if hit:
                    ls = hit["to"]
                    target = hit.get("detail_to")
                    if target:
                        if out.get(target) and out[target] != sub:
                            self._flag(rec, "CONFLICTING_SOURCE_VALUES", target, "Translated v1 detail conflicts with an existing value")
                        out[target] = sub
            if ls:
                self._flag(rec, "LEAD_SOURCE_TRANSLATED", "lead_source", f"v1 value translated to '{ls}'")
            else:
                self._flag(rec, "UNMAPPED_VALUE", "lead_source", f"Lead Source '{raw}' with detail '{details.get(rule.get('by_detail', ''))}' has no approved mapping")
        else:
            self._flag(rec, "UNMAPPED_VALUE", "lead_source", f"Lead Source '{raw}' has no approved mapping")
        out["lead_source"] = ls if ls else raw  # held records keep the received value for review

        if details.get("digital_v1") and raw != "Digital":
            self._flag(rec, "UNMAPPED_COLUMN", "digital_v1", "v1 Digital detail present without Lead Source 'Digital'")
        if details.get("referral_v1") and raw != "Referral":
            self._flag(rec, "UNMAPPED_COLUMN", "referral_v1", "v1 Referral detail present without Lead Source 'Referral'")

        if ls:
            paired = cfg.profile.lead_source_detail.get(ls)
            allowed = {paired} | set((cfg.profile.raw.get("lead_source_secondary") or {}).get(ls, []))
            for fid in _DEST_DETAILS:
                if out.get(fid) and fid not in allowed:
                    self._flag(rec, "ORPHAN_DETAIL", fid, f"Detail filled but Lead Source is '{ls}'")

    def _awareness(self, rec: Record, out: dict, v: dict) -> None:
        raw = self._text(rec, v.get(self.cmap.response["awareness"]), "awareness")
        if raw is None:
            return  # blank stays blank (v1 rows have no awareness question)
        tokens, unknown = tokenize_options(str(raw), self.cfg.vocab.awareness)
        for u in unknown:
            self._flag(rec, "UNKNOWN_OPTION", "awareness", f"Unknown awareness option: {u}")
        out["awareness"] = join_options(tokens)

    def _others(self, rec: Record, out: dict, v: dict) -> None:
        details = rec.meta.get("details_raw", {})
        notes = []
        for sid, col in self.cmap.others.items():
            spec = self.others_specs[sid]
            text = self._text(rec, v.get(col), sid)
            if text is None:
                continue
            parent = rec.lead_source_original if spec["parent"] == "lead_source" else details.get(spec["parent"])
            parent_tokens = [p.strip() for p in str(parent or "").split(",")]
            if "Others" not in parent_tokens:
                self._flag(rec, "ORPHAN_OTHERS_TEXT", sid, f"'Others' text present but '{spec['label']}' is not 'Others'")
            notes.append(f"[{spec['label']} – Others] {text}")
        if notes:
            out["remarks"] = " | ".join(([str(out["remarks"])] if out.get("remarks") else []) + notes)

    def _keys(self, rec: Record, submitter: object) -> None:
        out = rec.values
        ts = rec.timestamp.isoformat(timespec="microseconds") if rec.timestamp else ""
        digits = re.sub(r"\D", "", str(out.get("mobile") or ""))
        rec.src_key = _sha("FUEL-MARF", ts, str(submitter or "").lower(), rec.meta.get("seller_name") or "",
                           str(out.get("surname") or "").casefold(), digits)
        skip = {"id", "start_time", "completion_time"} | {f"seller_{k}_{n}" for k in ("name", "email") for n in (1, 2, 3)}
        content = {k: (str(val) if val is not None else None) for k, val in out.items() if k not in skip}
        content["_seller_name"] = rec.meta.get("seller_name")
        content["_seller_email"] = rec.meta.get("seller_email")
        rec.content_hash = hashlib.sha256(json.dumps(content, sort_keys=True).encode("utf-8")).hexdigest()
        if digits and out.get("surname") and out.get("first_name"):
            rec.client_key = _sha(digits[-10:], str(out["surname"]).casefold(), str(out["first_name"]).casefold())
