"""Field-level transformation and validation rules (design C, D.6, D.8).

Every function preserves the received value. The only change applied is the approved
leading/trailing whitespace trim (H-25). Nothing is corrected or invented.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime

_EMAIL = re.compile(r"^[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$")
_MOBILE_OK = re.compile(r"^(09\d{9}|\+639\d{9})$")


@dataclass
class Cleaned:
    value: object
    trimmed: bool = False
    null_token: bool = False


def clean_text(raw: object, null_tokens: set[str] | None = None) -> Cleaned:
    """Trim a text value; blank approved null tokens (H-11). Non-text values pass through."""
    if raw is None:
        return Cleaned(None)
    if isinstance(raw, bool):
        return Cleaned(str(raw))
    if isinstance(raw, int):
        return Cleaned(str(raw))
    if isinstance(raw, float):
        # Digits typed into a text question arrive as numeric cells; keep them as typed (no ".0").
        return Cleaned(str(int(raw)) if raw.is_integer() else repr(raw))
    if not isinstance(raw, str):
        return Cleaned(raw)
    stripped = raw.strip()
    trimmed = stripped != raw
    if stripped == "":
        return Cleaned(None, trimmed)
    if null_tokens and stripped in null_tokens:
        return Cleaned(None, trimmed, True)
    return Cleaned(stripped, trimmed)


def is_valid_email(value: str) -> bool:
    """Syntax only (H-24): one @, no spaces or commas, dotted domain, valid characters."""
    return bool(_EMAIL.match(value))


@dataclass
class MobileResult:
    text: str | None
    was_numeric: bool
    reasons: list[str]


def classify_mobile(raw: object) -> MobileResult:
    """Return the mobile as text exactly as received and any format reasons (D.8).

    A numeric cell is written as its digits; no leading 0 is ever added.
    """
    if raw is None or (isinstance(raw, str) and raw.strip() == ""):
        return MobileResult(None, False, [])
    was_numeric = isinstance(raw, (int, float)) and not isinstance(raw, bool)
    if was_numeric:
        text = str(int(raw)) if float(raw).is_integer() else repr(raw)
    else:
        text = str(raw).strip()
    if _MOBILE_OK.match(text):
        return MobileResult(text, was_numeric, [])
    reasons = []
    digits = re.sub(r"\D", "", text)
    if re.search(r"[^\d+]", text) or (text.count("+") and not text.startswith("+")):
        reasons.append("non_digit_characters")
    if len(digits) == 10 and digits.startswith("9"):
        reasons.append("missing_leading_zero")
    elif not (len(digits) == 11 and digits.startswith("09")) and not (len(digits) == 12 and digits.startswith("639")):
        reasons.append("wrong_length")
    if len(digits) >= 7 and len(set(digits)) == 1:
        reasons.append("repeated_digits")
    return MobileResult(text, was_numeric, reasons or ["wrong_length"])


def tokenize_options(raw: str, known: list[str]) -> tuple[list[str], list[str]]:
    """Split a multi-select answer into options (X2, R4).

    Google exports join options with ", " and some options contain commas, so the known
    options are matched longest-first. A value already joined with ";" is split on ";".
    Returns (tokens, unknown_tokens). Unknown text is kept as its own token.
    """
    text = raw.strip()
    if ";" in text:
        parts = [p.strip() for p in text.split(";") if p.strip()]
        return parts, [p for p in parts if p not in known]
    ordered = sorted(known, key=len, reverse=True)
    tokens, unknown = [], []
    i = 0
    while i < len(text):
        while i < len(text) and text[i] in ", ":
            i += 1
        if i >= len(text):
            break
        for opt in ordered:
            end = i + len(opt)
            if text.startswith(opt, i) and (end == len(text) or text.startswith(",", end)):
                tokens.append(opt)
                i = end
                break
        else:
            end = text.find(", ", i)
            end = len(text) if end == -1 else end
            piece = text[i:end].strip()
            tokens.append(piece)
            unknown.append(piece)
            i = end
    return tokens, unknown


def join_options(tokens: list[str]) -> str:
    return ";".join(tokens)


_AUTO_FORMATS = ("%m/%d/%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%d/%m/%Y %H:%M")
_ISO_FORMATS = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M")


@dataclass
class TimestampResult:
    value: datetime | None
    code: str | None  # MISSING_TIMESTAMP | INVALID_DATE | AMBIGUOUS_DATE
    message: str = ""


def parse_timestamp(raw: object, formats: list[str] | None = None, now: datetime | None = None) -> TimestampResult:
    """Keep datetimes unchanged; parse text only with confirmed or unambiguous formats (T14, T37)."""
    if raw is None or (isinstance(raw, str) and raw.strip() == ""):
        return TimestampResult(None, "MISSING_TIMESTAMP", "Timestamp is blank")
    if isinstance(raw, datetime):
        value = raw
    else:
        text = str(raw).strip()
        value = None
        for fmt in _ISO_FORMATS:
            try:
                value = datetime.strptime(text, fmt)
                break
            except ValueError:
                pass
        if value is None:
            candidates = {}
            for fmt in (formats or _AUTO_FORMATS):
                try:
                    candidates[fmt] = datetime.strptime(text, fmt)
                except ValueError:
                    pass
            distinct = set(candidates.values())
            if not distinct:
                return TimestampResult(None, "INVALID_DATE", "Timestamp text could not be read")
            if len(distinct) > 1:
                return TimestampResult(None, "AMBIGUOUS_DATE", "Day/month order is ambiguous; confirm the timestamp format")
            value = distinct.pop()
    if now is not None and value > now:
        return TimestampResult(None, "INVALID_DATE", "Timestamp is in the future")
    return TimestampResult(value, None)
