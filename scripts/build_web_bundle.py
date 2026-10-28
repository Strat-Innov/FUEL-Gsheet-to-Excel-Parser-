"""Bundle the parser source and shipped config for the browser app (public/fuel_bundle.json).

Run after changing fuel_parser/, config/ or web/bridge.py:  python scripts/build_web_bundle.py
tests/test_web.py fails if the bundle is out of date.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "public" / "fuel_bundle.json"
EXCLUDE = {"config/sellers.yaml", "config/sellers.csv", "config/teams.local.yaml"}


def collect() -> dict[str, str]:
    files = {}
    for pattern in ("fuel_parser/*.py", "config/*.yaml"):
        for p in sorted(ROOT.glob(pattern)):
            rel = p.relative_to(ROOT).as_posix()
            if rel not in EXCLUDE:
                files[rel] = p.read_text(encoding="utf-8")
    files["bridge.py"] = (ROOT / "web" / "bridge.py").read_text(encoding="utf-8")
    return files


def render() -> str:
    files = collect()
    digest = hashlib.sha256(json.dumps(files, sort_keys=True).encode("utf-8")).hexdigest()[:12]
    return json.dumps({"version": digest, "files": files}, indent=1, sort_keys=True) + "\n"


if __name__ == "__main__":
    OUT.write_text(render(), encoding="utf-8")
    print(f"wrote {OUT.relative_to(ROOT)}")
