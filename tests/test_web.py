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


# -- browser app (H-32) -----------------------------------------------------------------------

PUBLIC = ROOT / "public"


def test_web_bundle_is_up_to_date():
    import runpy
    build = runpy.run_path(str(ROOT / "scripts" / "build_web_bundle.py"))
    assert (PUBLIC / "fuel_bundle.json").read_text(encoding="utf-8") == build["render"](), \
        "public/fuel_bundle.json is stale: run python scripts/build_web_bundle.py"


def test_bundle_excludes_personal_data_files():
    import json
    files = json.loads((PUBLIC / "fuel_bundle.json").read_text(encoding="utf-8"))["files"]
    assert not any(name.startswith("config/sellers") and not name.endswith("example.yaml") for name in files)
    assert "config/teams.yaml" in files and "bridge.py" in files


def test_browser_code_only_talks_to_its_own_site_and_the_pyodide_cdn():
    import re
    import json
    text = "".join((PUBLIC / n).read_text(encoding="utf-8") for n in ("index.html", "app.js", "worker.js"))
    hosts = set(re.findall(r"https?://([a-z0-9.-]+)", text))
    assert hosts <= {"cdn.jsdelivr.net", "openapi.vercel.sh"}, hosts
    csp = json.loads((ROOT / "vercel.json").read_text())["headers"][0]["headers"][0]["value"]
    assert "connect-src 'self' https://cdn.jsdelivr.net;" in csp and "form-action 'none'" in csp
    assert "XMLHttpRequest" not in text and "sendBeacon" not in text


def test_browser_ui_inserts_file_values_as_text_only():
    js = (PUBLIC / "app.js").read_text(encoding="utf-8")
    assert "innerHTML" not in js and "insertAdjacentHTML" not in js and "document.write" not in js


def test_vendored_wheels_match_pinned_hashes():
    import hashlib
    pins = {"openpyxl-3.1.5-py2.py3-none-any.whl": "5282c12b107bffeef825f4617dc029afaf41d0ea60823bbb665ef3079dc79de2",
            "et_xmlfile-2.0.0-py3-none-any.whl": "7a91720bc756843502c3b7504c77b8fe44217c85c537d85037f0f536151b2caa"}
    for name, digest in pins.items():
        assert hashlib.sha256((PUBLIC / "wheels" / name).read_bytes()).hexdigest() == digest


def _bridge():
    import sys
    sys.path.insert(0, str(ROOT / "web"))
    import bridge
    bridge.APP = ROOT
    return bridge


def test_classify_reads_headers_not_names(tmp_path):
    import json
    import openpyxl
    from tests.fixtures.builders import TEMPLATE_HEADERS, response, write_source, write_template
    b = _bridge()
    write_source(tmp_path / "Sales FUEL Template.xlsx", [response()])  # misleading file name
    write_template(tmp_path / "responses.xlsx")
    wb = openpyxl.Workbook(); wb.active.title = "Sheet1"; wb.active.append(TEMPLATE_HEADERS); wb.save(tmp_path / "renamed.xlsx")
    wb = openpyxl.Workbook(); wb.active.append(TEMPLATE_HEADERS[:-1] + ["changed"]); wb.save(tmp_path / "almost.xlsx")
    (tmp_path / "dir.csv").write_text("seller_name,proposed_email\n", encoding="utf-8")
    got = {n: json.loads(b.classify(str(tmp_path / n))) for n in ("Sales FUEL Template.xlsx", "responses.xlsx", "renamed.xlsx", "almost.xlsx", "dir.csv")}
    assert got["Sales FUEL Template.xlsx"]["kind"] == "source"
    assert got["responses.xlsx"]["kind"] == "template" and got["responses.xlsx"]["template_exact"]
    assert got["renamed.xlsx"]["template_sheet"] == "Sheet1" and got["renamed.xlsx"]["template_exact"]
    assert got["almost.xlsx"]["kind"] == "template" and not got["almost.xlsx"]["template_exact"]
    assert got["dir.csv"]["kind"] == "sellers"
