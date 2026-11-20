"""Check a COBie workbook (xlsx) for missing sheets, columns, bad references and bad values."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime

from openpyxl import load_workbook

COMMON = ["Name", "CreatedBy", "CreatedOn"]
REQUIRED_SHEETS = {
    "Contact": ["Email", "CreatedBy", "CreatedOn"],
    "Facility": COMMON,
    "Floor": COMMON,
    "Space": COMMON + ["FloorName"],
    "Type": COMMON,
    "Component": COMMON + ["TypeName", "Space"],
}
ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}")


@dataclass(frozen=True)
class Issue:
    sheet: str
    row: int  # 0 means the whole sheet or the header
    message: str

    def __str__(self) -> str:
        where = f"{self.sheet} row {self.row}" if self.row else self.sheet
        return f"{where}: {self.message}"


def _text(v) -> str:
    return "" if v is None else str(v).strip()


def read_sheets(path: str) -> dict[str, tuple[list[str], list[tuple[int, dict]]]]:
    """Return {sheet: (header, [(excel row number, {column: value})])}. Blank rows are skipped."""
    wb = load_workbook(path, read_only=True, data_only=True)
    out = {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            out[ws.title] = ([], [])
            continue
        header = [_text(h) for h in rows[0]]
        body = []
        for i, r in enumerate(rows[1:], start=2):
            if all(_text(v) == "" for v in r):
                continue
            body.append((i, {h: r[j] if j < len(r) else None for j, h in enumerate(header) if h}))
        out[ws.title] = (header, body)
    return out


def _valid_date(v) -> bool:
    if isinstance(v, (datetime, date)):
        return True
    return bool(ISO_DATE.match(_text(v)))


def validate_workbook(path: str) -> list[Issue]:
    sheets = read_sheets(path)
    issues: list[Issue] = []
    for name, cols in REQUIRED_SHEETS.items():
        if name not in sheets:
            issues.append(Issue(name, 0, "required sheet is missing"))
            continue
        header = sheets[name][0]
        for c in cols:
            if c not in header:
                issues.append(Issue(name, 0, f"required column '{c}' is missing"))

    def rows(name):
        return sheets.get(name, ([], []))[1]

    def names(sheet, col="Name"):
        return {_text(r.get(col)) for _, r in rows(sheet) if _text(r.get(col))}

    contacts = {_text(r.get("Email")).lower() for _, r in rows("Contact")}
    floors, spaces, types = names("Floor"), names("Space"), names("Type")

    for name in REQUIRED_SHEETS:
        if name not in sheets:
            continue
        key = "Email" if name == "Contact" else "Name"
        seen: dict[str, int] = {}
        for num, r in rows(name):
            ident = _text(r.get(key))
            if key in sheets[name][0] and not ident:
                issues.append(Issue(name, num, f"{key} is empty"))
            elif ident:
                if ident.lower() in seen:
                    issues.append(Issue(name, num, f"duplicate {key} '{ident}' (first on row {seen[ident.lower()]})"))
                seen.setdefault(ident.lower(), num)
            if "CreatedBy" in r:
                who = _text(r["CreatedBy"]).lower()
                if not who:
                    issues.append(Issue(name, num, "CreatedBy is empty"))
                elif contacts and who not in contacts:
                    issues.append(Issue(name, num, f"CreatedBy '{_text(r['CreatedBy'])}' is not in the Contact sheet"))
            if "CreatedOn" in r and not _valid_date(r["CreatedOn"]):
                issues.append(Issue(name, num, f"CreatedOn '{_text(r['CreatedOn'])}' is not a YYYY-MM-DD date"))

    for num, r in rows("Space"):
        if "FloorName" in r and _text(r["FloorName"]) not in floors:
            issues.append(Issue("Space", num, f"FloorName '{_text(r['FloorName'])}' is not in the Floor sheet"))
        area = r.get("GrossArea")
        if _text(area):
            try:
                if float(area) <= 0:
                    issues.append(Issue("Space", num, f"GrossArea {area} must be greater than 0"))
            except (TypeError, ValueError):
                issues.append(Issue("Space", num, f"GrossArea '{_text(area)}' is not a number"))
    for num, r in rows("Component"):
        if "TypeName" in r and _text(r["TypeName"]) not in types:
            issues.append(Issue("Component", num, f"TypeName '{_text(r['TypeName'])}' is not in the Type sheet"))
        if "Space" in r:
            sp = _text(r["Space"])
            if not sp:
                issues.append(Issue("Component", num, "Space is empty"))
            elif sp not in spaces:
                issues.append(Issue("Component", num, f"Space '{sp}' is not in the Space sheet"))
        inst = r.get("InstallationDate")
        if _text(inst) and not _valid_date(inst):
            issues.append(Issue("Component", num, f"InstallationDate '{_text(inst)}' is not a YYYY-MM-DD date"))
    return issues
