"""Unit tests for header matching, transforms and configuration validation."""

from datetime import datetime
from pathlib import Path

import pytest

from fuel_parser.config import ConfigError, load_teams
from fuel_parser.headers import match_key
from fuel_parser.transform import (classify_mobile, clean_text, is_valid_email, parse_timestamp,
                                   tokenize_options)

AWARENESS = ["Word of Mouth", "Social Media (Facebook, Instagram, etc)", "Events"]


def test_match_key_normalises_nbsp_suffix_colon_and_case():
    assert match_key("If\xa0Booth, please choose which\xa0Booth Location\xa0below:\n") == \
        "if booth, please choose which booth location below"
    assert match_key("Seller Name  13") == "seller name"
    assert match_key("Client's Middle Name ") == "client's middle name"
    assert match_key("Seller Name1") == "seller name1"  # MS Forms suffix has no space: kept distinct


def test_T10_awareness_with_commas_inside_options():
    tokens, unknown = tokenize_options("Word of Mouth, Social Media (Facebook, Instagram, etc), Events", AWARENESS)
    assert tokens == ["Word of Mouth", "Social Media (Facebook, Instagram, etc)", "Events"]
    assert unknown == []


def test_T11_unknown_awareness_token_is_kept():
    tokens, unknown = tokenize_options("Word of Mouth, Radio", AWARENESS)
    assert tokens == ["Word of Mouth", "Radio"] and unknown == ["Radio"]


def test_semicolon_input_is_split_on_semicolons():
    tokens, unknown = tokenize_options("Events;Word of Mouth", AWARENESS)
    assert tokens == ["Events", "Word of Mouth"] and unknown == []


@pytest.mark.parametrize("raw,text,numeric,reasons", [
    ("09000001234", "09000001234", False, []),                                   # T15
    ("+639000001234", "+639000001234", False, []),
    (9000001234, "9000001234", True, ["missing_leading_zero"]),                  # T16
    ("9999999999", "9999999999", False, ["missing_leading_zero", "repeated_digits"]),  # T31
    ("0900-000-1234", "0900-000-1234", False, ["non_digit_characters"]),
    ("12345", "12345", False, ["wrong_length"]),
    (None, None, False, []),
])
def test_mobile_is_kept_exactly_and_classified(raw, text, numeric, reasons):
    m = classify_mobile(raw)
    assert (m.text, m.was_numeric, m.reasons) == (text, numeric, reasons)


@pytest.mark.parametrize("value,ok", [
    ("name@gmail,com", False), ("name @x.com", False), ("a@b", False), ("a@@b.com", False),
    ("name@gmaol.com", True), ("x@company.ph", True), ("First.Last+tag@example.co.uk", True),  # T33
])
def test_email_syntax_only(value, ok):
    assert is_valid_email(value) is ok


def test_T14_dates():
    assert parse_timestamp(None).code == "MISSING_TIMESTAMP"
    assert parse_timestamp("yesterday").code == "INVALID_DATE"
    assert parse_timestamp("03/04/2026 10:00:00").code == "AMBIGUOUS_DATE"
    assert parse_timestamp("03/04/2026 10:00:00", ["%d/%m/%Y %H:%M:%S"]).value == datetime(2026, 4, 3, 10, 0)
    assert parse_timestamp("03/14/2026 10:00:00").value == datetime(2026, 3, 14, 10, 0)
    assert parse_timestamp(datetime(2030, 1, 1), now=datetime(2026, 1, 1)).code == "INVALID_DATE"
    keep = datetime(2026, 5, 6, 21, 5, 48, 239000)
    assert parse_timestamp(keep).value == keep  # milliseconds preserved


def test_clean_text_trims_blanks_null_tokens_and_keeps_numeric_text():
    assert clean_text("  Joe ").value == "Joe" and clean_text("  Joe ").trimmed
    assert clean_text("N/A", {"N/A"}).null_token
    assert clean_text(9000001234.0).value == "9000001234"  # no ".0" for digits typed in a text field
    assert clean_text("   ").value is None


def _teams(tmp_path: Path, body: str):
    p = tmp_path / "teams.yaml"
    p.write_text("teams:\n" + body, encoding="utf-8")
    return load_teams(p)


@pytest.mark.parametrize("body,msg", [
    ('  - {label_prefix: "T (AAD:", slot: 1, status: VERIFIED, evidence_type: masterlist_export}\n', "evidence_ref"),
    ('  - {label_prefix: "T (AAD:", slot: 1, status: VERIFIED, evidence_type: business_rule, evidence_ref: x}\n', "evidence_type"),
    ('  - {label_prefix: "T (AAD:", slot: 1, status: OWNER_APPROVED_UNVERIFIED, approved_by: me}\n', "approved_on"),
    ('  - {label_prefix: "T (AAD:", slot: 2, status: UNRESOLVED}\n', "slot: null"),
    ('  - {label_prefix: "T (AAD:", slot: null, status: VERIFIED, evidence_type: masterlist_export, evidence_ref: x}\n', "requires a slot"),
    ('  - {label_prefix: "T (AAD:", slot: 4, status: UNRESOLVED}\n', "slot must be"),
    ('  - {label_prefix: "T (AAD:", slot: null, status: MAYBE}\n', "status must be"),
])
def test_T47_teams_config_rejected(tmp_path, body, msg):
    with pytest.raises(ConfigError, match=msg.replace("(", r"\(")):
        _teams(tmp_path, body)


def test_shipped_teams_config_has_every_team_unresolved():
    repo = Path(__file__).resolve().parent.parent
    teams = load_teams(repo / "config" / "teams.yaml")
    assert len(teams.entries) == 8
    assert all(e.status == "UNRESOLVED" and e.slot is None and not e.exportable for e in teams.entries)


def test_T49_parser_code_contains_no_team_labels_or_slots():
    repo = Path(__file__).resolve().parent.parent
    teams = load_teams(repo / "config" / "teams.yaml")
    code = "\n".join(p.read_text(encoding="utf-8") for p in (repo / "fuel_parser").glob("*.py"))
    code += (repo / "app" / "streamlit_app.py").read_text(encoding="utf-8")
    for e in teams.entries:
        team_word = e.label_prefix.split("(")[0].strip()
        assert team_word not in code, f"team label '{team_word}' found in parser code"
