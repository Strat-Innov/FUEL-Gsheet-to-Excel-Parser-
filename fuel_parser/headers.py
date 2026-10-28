"""Header normalisation and alias matching (design section C).

Headers are matched by meaning through a normalised *match key*, never by
column letter.
"""

from __future__ import annotations

import difflib
import re
import unicodedata

_SUFFIX = re.compile(r"\s+\d+$")
_SPACE = re.compile(r"\s+")


def match_key(header: object) -> str:
    """Return the normalised match key for a header cell.

    NFKC -> NBSP to space -> lower-case -> collapse whitespace -> strip ->
    drop the Google duplicate suffix (`` 2``, `` 13``) -> drop a trailing ``:``.
    """
    if header is None:
        return ""
    text = unicodedata.normalize("NFKC", str(header)).replace(" ", " ")
    text = _SPACE.sub(" ", text.lower()).strip()
    text = _SUFFIX.sub("", text).strip()
    return text.rstrip(":").strip()


def suggest(key: str, candidates: list[str], cutoff: float = 0.85) -> str | None:
    """Closest alias for an unmatched key. A suggestion is never auto-accepted."""
    found = difflib.get_close_matches(key, candidates, n=1, cutoff=cutoff)
    return found[0] if found else None
