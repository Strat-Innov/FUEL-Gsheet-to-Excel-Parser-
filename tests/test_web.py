"""The Vercel entrypoint is information-only (H-31)."""

import importlib.util
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("vercel_index", ROOT / "api" / "index.py")
index = importlib.util.module_from_spec(spec)
spec.loader.exec_module(index)


def call(method):
    out = {}
    body = b"".join(index.app({"REQUEST_METHOD": method, "PATH_INFO": "/"},
                              lambda status, headers: out.update(status=status, headers=dict(headers))))
    return out["status"], out["headers"], body


def test_get_returns_static_page():
    status, headers, body = call("GET")
    assert status.startswith("200") and headers["Content-Type"].startswith("text/html")
    assert b"FUEL Response Parser" in body and b"form-action 'none'" in headers["Content-Security-Policy"].encode()


def test_page_has_no_upload_or_form():
    _, _, body = call("GET")
    assert b"<form" not in body.lower() and b"<input" not in body.lower()


def test_post_is_refused():
    status, _, _ = call("POST")
    assert status.startswith("405")


def test_entrypoint_does_not_import_the_parser():
    src = (ROOT / "api" / "index.py").read_text(encoding="utf-8")
    assert "import fuel_parser" not in src and "from fuel_parser" not in src
