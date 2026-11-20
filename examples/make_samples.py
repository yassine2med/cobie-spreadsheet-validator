"""Write the two sample COBie workbooks in this folder (invented data)."""

from pathlib import Path

from openpyxl import Workbook

HERE = Path(__file__).resolve().parent
EMAIL = "fm-contact"

SHEETS = {
    "Contact": (["Email", "CreatedBy", "CreatedOn", "Category", "Company"],
                [[EMAIL, EMAIL, "2026-01-15", "Facility Manager", "Sample Facilities Ltd"],
                 ["design-contact", EMAIL, "2026-01-15", "Architect", "Sample Design Studio"]]),
    "Facility": (["Name", "CreatedBy", "CreatedOn", "Category", "ProjectName", "SiteName", "LinearUnits", "AreaUnits"],
                 [["Harbor Office", EMAIL, "2026-01-15", "Office", "Harbor Fit-out", "Harbor Campus", "meters", "square meters"]]),
    "Floor": (["Name", "CreatedBy", "CreatedOn", "Category", "Elevation"],
              [["Level 1", EMAIL, "2026-01-15", "Floor", 0], ["Level 2", EMAIL, "2026-01-15", "Floor", 3.5]]),
    "Space": (["Name", "CreatedBy", "CreatedOn", "Category", "FloorName", "GrossArea"],
              [["101", EMAIL, "2026-01-15", "Office", "Level 1", 42.5],
               ["102", EMAIL, "2026-01-15", "Meeting", "Level 1", 18.0],
               ["201", EMAIL, "2026-01-15", "Plant", "Level 2", 25.25]]),
    "Type": (["Name", "CreatedBy", "CreatedOn", "Category", "Manufacturer", "WarrantyDurationParts"],
             [["Split Heat Pump", EMAIL, "2026-01-15", "HVAC", "Acme Climate", 5],
              ["LED Panel", EMAIL, "2026-01-15", "Lighting", "Bright Co", 3]]),
    "Component": (["Name", "CreatedBy", "CreatedOn", "TypeName", "Space", "SerialNumber", "InstallationDate"],
                  [["HP-101", EMAIL, "2026-01-15", "Split Heat Pump", "101", "SN-1001", "2025-11-03"],
                   ["HP-102", EMAIL, "2026-01-15", "Split Heat Pump", "102", "SN-1002", "2025-11-03"],
                   ["LP-201", EMAIL, "2026-01-15", "LED Panel", "201", "SN-2001", "2025-12-01"]]),
}


def build(path: Path, mutate=None) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    data = {k: (list(h), [list(r) for r in rows]) for k, (h, rows) in SHEETS.items()}
    if mutate:
        mutate(data)
    for name, (head, rows) in data.items():
        ws = wb.create_sheet(name)
        ws.append(head)
        for r in rows:
            ws.append(r)
    wb.save(path)


def break_it(data):
    data["Space"][1].append(["101", EMAIL, "2026-01-15", "Office", "Level 1", 30])      # duplicate name
    data["Space"][1].append(["301", EMAIL, "15/01/2026", "Office", "Level 9", -4])      # bad date, floor, area
    data["Component"][1][1][3] = "Split Heat Pumps"                                    # unknown type
    data["Component"][1][2][4] = ""                                                   # missing space
    data["Type"][1][1][1] = "nobody-contact"                                       # unknown contact
    del data["Floor"][0][2]
    for r in data["Floor"][1]:
        del r[2]


if __name__ == "__main__":
    build(HERE / "cobie_valid.xlsx")
    build(HERE / "cobie_with_errors.xlsx", break_it)
    print("wrote cobie_valid.xlsx and cobie_with_errors.xlsx")
