"""Serve the browser app locally: python scripts/serve_web.py [--port 8000] [--pyodide DIR]

Uses the same security headers as vercel.json. Pyodide loads from cdn.jsdelivr.net unless --pyodide
points to a local copy of the Pyodide distribution (served under /pyodide/).
"""

from __future__ import annotations

import argparse
import functools
import json
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HEADERS = {h["key"]: h["value"] for h in json.loads((ROOT / "vercel.json").read_text())["headers"][0]["headers"]}


class Handler(SimpleHTTPRequestHandler):
    pyodide_dir: Path | None = None

    def translate_path(self, path: str) -> str:
        clean = path.split("?", 1)[0].split("#", 1)[0]
        if clean.startswith("/pyodide/") and self.pyodide_dir:
            return str(self.pyodide_dir / clean[len("/pyodide/"):])
        return super().translate_path(path)

    def end_headers(self) -> None:
        for k, v in HEADERS.items():
            self.send_header(k, v)
        super().end_headers()

    def log_message(self, *args) -> None:  # keep the console quiet
        pass


Handler.extensions_map = {**SimpleHTTPRequestHandler.extensions_map, ".wasm": "application/wasm", ".js": "text/javascript",
                          ".mjs": "text/javascript", ".json": "application/json", ".whl": "application/zip"}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8000)
    ap.add_argument("--pyodide", type=Path, help="Local Pyodide distribution directory (optional)")
    a = ap.parse_args()
    Handler.pyodide_dir = a.pyodide
    server = ThreadingHTTPServer(("127.0.0.1", a.port), functools.partial(Handler, directory=str(ROOT / "public")))
    print(f"FUEL Response Parser: http://localhost:{a.port}  (Ctrl+C to stop)")
    server.serve_forever()


if __name__ == "__main__":
    main()
