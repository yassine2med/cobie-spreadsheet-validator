"""Command line: python -m cobievalidator check FILE.xlsx"""

from __future__ import annotations

import argparse
import sys

from .validator import read_sheets, validate_workbook


def main(argv: list[str] | None = None) -> int:
    """Return 0 when clean, 1 when issues exist, 2 when the file cannot be read."""
    parser = argparse.ArgumentParser(prog="cobievalidator", description="COBie workbook validator")
    sub = parser.add_subparsers(dest="command", required=True)
    chk = sub.add_parser("check", help="list every problem found in a COBie workbook")
    chk.add_argument("file")
    sm = sub.add_parser("sheets", help="show each sheet with its row count")
    sm.add_argument("file")
    args = parser.parse_args(argv)
    try:
        if args.command == "sheets":
            for name, (header, rows) in read_sheets(args.file).items():
                print(f"{name:<12} {len(rows):>3} rows, {len(header)} columns")
            return 0
        issues = validate_workbook(args.file)
    except FileNotFoundError:
        print(f"File not found: {args.file}", file=sys.stderr)
        return 2
    except Exception as exc:  # not an xlsx file, corrupt zip, and so on
        print(f"Cannot read {args.file}: {exc}", file=sys.stderr)
        return 2
    for issue in issues:
        print(issue)
    print(f"{len(issues)} issue(s)")
    return 1 if issues else 0


if __name__ == "__main__":
    sys.exit(main())
