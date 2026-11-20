import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "examples"))

import make_samples  # noqa: E402
from cobievalidator.__main__ import main  # noqa: E402
from cobievalidator.validator import read_sheets, validate_workbook  # noqa: E402


def check(tmp_path, mutate=None):
    p = tmp_path / "t.xlsx"
    make_samples.build(p, mutate)
    return [str(i) for i in validate_workbook(str(p))]


def test_valid_sample_has_no_issues():
    assert validate_workbook(str(ROOT / "examples" / "cobie_valid.xlsx")) == []


def test_sample_with_errors_has_issues():
    issues = [str(i) for i in validate_workbook(str(ROOT / "examples" / "cobie_with_errors.xlsx"))]
    assert len(issues) == 8


def test_clean_built_workbook(tmp_path):
    assert check(tmp_path) == []


def test_missing_sheet(tmp_path):
    def m(d):
        del d["Type"]
    assert "Type: required sheet is missing" in check(tmp_path, m)


def test_missing_column(tmp_path):
    def m(d):
        i = d["Space"][0].index("FloorName")
        del d["Space"][0][i]
        for r in d["Space"][1]:
            del r[i]
    assert "Space: required column 'FloorName' is missing" in check(tmp_path, m)


def test_duplicate_name(tmp_path):
    def m(d):
        d["Floor"][1].append(["Level 1", "fm-contact", "2026-01-15", "Floor", 7])
    assert any("duplicate Name 'Level 1' (first on row 2)" in i for i in check(tmp_path, m))


def test_duplicate_is_case_insensitive(tmp_path):
    def m(d):
        d["Floor"][1].append(["LEVEL 1", "fm-contact", "2026-01-15", "Floor", 7])
    assert any("duplicate" in i for i in check(tmp_path, m))


def test_unknown_floor(tmp_path):
    def m(d):
        d["Space"][1][0][4] = "Level 7"
    assert "Space row 2: FloorName 'Level 7' is not in the Floor sheet" in check(tmp_path, m)


def test_unknown_type_and_space(tmp_path):
    def m(d):
        d["Component"][1][0][3] = "Nothing"
        d["Component"][1][1][4] = "999"
    out = check(tmp_path, m)
    assert any("TypeName 'Nothing'" in i for i in out) and any("Space '999'" in i for i in out)


def test_unknown_contact(tmp_path):
    def m(d):
        d["Floor"][1][0][1] = "ghost-contact"
    assert any("not in the Contact sheet" in i for i in check(tmp_path, m))


def test_bad_date_formats(tmp_path):
    def m(d):
        d["Space"][1][0][2] = "01/15/2026"
        d["Component"][1][0][6] = "yesterday"
    out = check(tmp_path, m)
    assert any("CreatedOn '01/15/2026'" in i for i in out) and any("InstallationDate 'yesterday'" in i for i in out)


def test_area_rules(tmp_path):
    def m(d):
        d["Space"][1][0][5] = 0
        d["Space"][1][1][5] = "big"
    out = check(tmp_path, m)
    assert any("must be greater than 0" in i for i in out) and any("is not a number" in i for i in out)


def test_blank_rows_skipped(tmp_path):
    def m(d):
        d["Floor"][1].append([None] * 5)
    assert check(tmp_path, m) == []


def test_read_sheets_counts():
    s = read_sheets(str(ROOT / "examples" / "cobie_valid.xlsx"))
    assert len(s["Component"][1]) == 3 and s["Space"][0][0] == "Name"


def test_cli_exit_codes(capsys):
    assert main(["check", str(ROOT / "examples" / "cobie_valid.xlsx")]) == 0
    assert main(["check", str(ROOT / "examples" / "cobie_with_errors.xlsx")]) == 1
    assert main(["check", "missing.xlsx"]) == 2
    assert main(["sheets", str(ROOT / "examples" / "cobie_valid.xlsx")]) == 0


def test_cli_not_a_workbook(tmp_path):
    bad = tmp_path / "x.xlsx"
    bad.write_text("not a zip")
    assert main(["check", str(bad)]) == 2
