"""Vercel entrypoint: a static, read-only information page (decision H-31).

The FUEL parser processes client data locally only (H-29). This page accepts no files, reads no
workbooks and imports nothing from `fuel_parser`, so no client data ever reaches Vercel.
"""

from __future__ import annotations

REPO_URL = "https://github.com/Strat-Innov/FUEL-Gsheet-to-Excel-Parser-"

PAGE = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<title>FUEL Response Parser</title>
<style>
  :root {{ --bg: #ffffff; --fg: #1d1d1f; --muted: #5f6368; --line: #e3e3e3; --code: #f5f5f5; }}
  @media (prefers-color-scheme: dark) {{ :root {{ --bg: #141414; --fg: #ececec; --muted: #a0a0a0; --line: #2c2c2c; --code: #1f1f1f; }} }}
  body {{ margin: 0; background: var(--bg); color: var(--fg); font: 16px/1.55 system-ui, -apple-system, "Segoe UI", sans-serif; }}
  main {{ max-width: 720px; margin: 0 auto; padding: 48px 16px 64px; }}
  h1 {{ font-size: 1.8rem; margin: 0 0 4px; }}
  h2 {{ font-size: 1.1rem; margin: 32px 0 8px; }}
  p.lead {{ color: var(--muted); margin-top: 0; }}
  pre {{ background: var(--code); border: 1px solid var(--line); border-radius: 6px; padding: 12px; overflow-x: auto; font-size: 0.9rem; }}
  .note {{ border-left: 3px solid var(--muted); padding: 4px 12px; color: var(--muted); }}
  a {{ color: inherit; }}
</style>
</head>
<body>
<main>
  <h1>FUEL Response Parser</h1>
  <p class="lead">Converts FUEL Google Forms response exports (.xlsx or .csv) into the 35-column MASTERLIST import
  format, with a full audit and review report.</p>

  <p class="note">This page is information only. The parser runs on your own computer: it accepts no uploads here,
  and client data is never sent to this site.</p>

  <h2>Install (once, Python 3.11+)</h2>
  <pre>git clone {REPO_URL}
cd FUEL-Gsheet-to-Excel-Parser-
python -m venv .venv
.venv/bin/pip install -e ".[ui,dev]"</pre>

  <h2>Validate an export (dry run)</h2>
  <pre>.venv/bin/fuel-parser run --source "&lt;response export&gt;.xlsx" --template "&lt;MASTERLIST template&gt;.xlsx" --dry-run</pre>
  <p>The audit workbook in <code>output/</code> accounts for every response: exported, held for review, duplicates
  held, already imported and excluded.</p>

  <h2>Local screen</h2>
  <pre>.venv/bin/streamlit run app/streamlit_app.py</pre>

  <p>Full instructions: <a href="{REPO_URL}">README and design document</a>.</p>
</main>
</body>
</html>
""".encode("utf-8")

_HEADERS = [
    ("Content-Type", "text/html; charset=utf-8"),
    ("Content-Length", str(len(PAGE))),
    ("Cache-Control", "public, max-age=300"),
    ("X-Content-Type-Options", "nosniff"),
    ("Referrer-Policy", "no-referrer"),
    ("Content-Security-Policy", "default-src 'none'; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'"),
]


def app(environ, start_response):
    """WSGI app: GET/HEAD return the page; every other method is refused (nothing can be uploaded)."""
    method = environ.get("REQUEST_METHOD", "GET")
    if method not in ("GET", "HEAD"):
        body = b"Method not allowed: this page accepts no data.\n"
        start_response("405 Method Not Allowed", [("Allow", "GET, HEAD"), ("Content-Type", "text/plain; charset=utf-8"),
                                                  ("Content-Length", str(len(body)))])
        return [body]
    start_response("200 OK", _HEADERS)
    return [b"" if method == "HEAD" else PAGE]
