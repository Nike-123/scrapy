#!/usr/bin/env python3
"""Build the verified-firms spreadsheet (CSV + XLSX) from results.json.

Reads:
    ip_verification/input_domains.csv  (original rows, preserved verbatim)
    ip_verification/results.json       (per-row verification findings)

Writes:
    ip_verification/output/verified_ip_firms.csv
    ip_verification/output/verified_ip_firms.xlsx  (needs openpyxl)

Validation (fails the build rather than emitting suspect data):
    * every input Row_ID has exactly one result row and vice versa;
    * IP_Status / US_Based_Status / Verification_Status use the task enums;
    * an Owner_Email is present only together with an Email_Source_URL;
    * no generic mailbox (info@, contact@, ...) sits in Owner_Email.

Usage:
    python3 ip_verification/build_output.py
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from collections import Counter
from pathlib import Path

HERE = Path(__file__).resolve().parent
INPUT_CSV = HERE / "input_domains.csv"
RESULTS_JSON = HERE / "results.json"
OUT_DIR = HERE / "output"
OUT_CSV = OUT_DIR / "verified_ip_firms.csv"
OUT_XLSX = OUT_DIR / "verified_ip_firms.xlsx"

ORIGINAL_COLUMNS = ["Row_ID", "Company_Name_Original", "Website_URL_Original", "Domain"]
REQUIRED_COLUMNS = [
    "Website_Status",
    "Verified_Official_Website",
    "US_Based_Status",
    "US_Evidence_URL",
    "IP_Status",
    "IP_Practice_Areas",
    "IP_Evidence_URL",
    "Owner_First_Name",
    "Owner_Last_Name",
    "Owner_Title",
    "Owner_Email",
    "Email_Source_URL",
    "General_Email",
    "Verification_Status",
    "Review_Notes",
]
ALL_COLUMNS = (
    ORIGINAL_COLUMNS
    + ["Company_Name_Verified"]
    + REQUIRED_COLUMNS
    + ["Checked_Date"]
)

IP_ENUM = {"IP", "NIP", "Needs Review"}
US_ENUM = {"US-Based", "Non-US", "Needs Review"}
VERIFICATION_ENUM = {"Verified", "Needs Review"}
GENERIC_LOCALPARTS = {
    "info", "contact", "mailbox", "office", "admin", "hello", "frontdesk", "mail",
}


def load_inputs() -> tuple[list[dict[str, str]], dict]:
    with INPUT_CSV.open(newline="", encoding="utf-8") as fh:
        inputs = list(csv.DictReader(fh))
    results = json.loads(RESULTS_JSON.read_text(encoding="utf-8"))
    return inputs, results


def merge_rows(inputs: list[dict[str, str]], results: dict) -> list[dict[str, str]]:
    by_id = {int(r["Row_ID"]): r for r in results["rows"]}
    input_ids = {int(r["Row_ID"]) for r in inputs}
    if input_ids != set(by_id):
        missing = sorted(input_ids - set(by_id))
        extra = sorted(set(by_id) - input_ids)
        raise SystemExit(f"Row_ID mismatch. Missing results: {missing}; extra: {extra}")
    checked = results.get("meta", {}).get("checked_date", "")
    merged: list[dict[str, str]] = []
    for original in inputs:
        finding = by_id[int(original["Row_ID"])]
        row: dict[str, str] = {}
        for col in ALL_COLUMNS:
            if col in ORIGINAL_COLUMNS:
                row[col] = original.get(col, "")
            elif col == "Checked_Date":
                # Per-row date wins (used by the combined 1-110 build).
                row[col] = str(finding.get(col, "") or checked)
            else:
                row[col] = str(finding.get(col, "") or "")
        merged.append(row)
    return merged


def validate(rows: list[dict[str, str]]) -> None:
    problems: list[str] = []
    for row in rows:
        rid = row["Row_ID"]
        if row["IP_Status"] not in IP_ENUM:
            problems.append(f"Row {rid}: bad IP_Status {row['IP_Status']!r}")
        if row["US_Based_Status"] not in US_ENUM:
            problems.append(f"Row {rid}: bad US_Based_Status {row['US_Based_Status']!r}")
        if row["Verification_Status"] not in VERIFICATION_ENUM:
            problems.append(f"Row {rid}: bad Verification_Status {row['Verification_Status']!r}")
        email = row["Owner_Email"].strip()
        if email:
            if not row["Email_Source_URL"].strip():
                problems.append(f"Row {rid}: Owner_Email without Email_Source_URL")
            local = email.split("@")[0].lower()
            if local in GENERIC_LOCALPARTS:
                problems.append(f"Row {rid}: generic mailbox in Owner_Email: {email}")
            if "@" not in email or " " in email:
                problems.append(f"Row {rid}: malformed Owner_Email: {email!r}")
        if row["Website_Status"].startswith("Inaccessible") and row["IP_Status"] == "NIP":
            problems.append(f"Row {rid}: inaccessible site must not be auto-NIP")
    if problems:
        raise SystemExit("VALIDATION FAILED:\n" + "\n".join(problems))


def write_csv(rows: list[dict[str, str]]) -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    with OUT_CSV.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=ALL_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def write_xlsx(rows: list[dict[str, str]], meta: dict) -> bool:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("openpyxl not installed - skipping XLSX (pip install openpyxl)", file=sys.stderr)
        return False

    shortlist = [
        r for r in rows if r["US_Based_Status"] == "US-Based" and r["IP_Status"] == "IP"
    ]
    counts = summarise(rows)

    wb = Workbook()

    def fill_sheet(ws, data, widths):
        header_font = Font(bold=True)
        for idx, col in enumerate(ALL_COLUMNS, start=1):
            cell = ws.cell(row=1, column=idx, value=col)
            cell.font = header_font
            ws.column_dimensions[get_column_letter(idx)].width = widths.get(col, 18)
        wrap = Alignment(wrap_text=True, vertical="top")
        for r_idx, row in enumerate(data, start=2):
            for c_idx, col in enumerate(ALL_COLUMNS, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=row[col])
                cell.alignment = wrap
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    widths = {
        "Row_ID": 8,
        "Company_Name_Original": 20,
        "Website_URL_Original": 34,
        "Domain": 30,
        "Company_Name_Verified": 34,
        "Website_Status": 20,
        "Verified_Official_Website": 36,
        "US_Based_Status": 14,
        "US_Evidence_URL": 36,
        "IP_Status": 12,
        "IP_Practice_Areas": 44,
        "IP_Evidence_URL": 36,
        "Owner_First_Name": 14,
        "Owner_Last_Name": 16,
        "Owner_Title": 34,
        "Owner_Email": 30,
        "Email_Source_URL": 36,
        "General_Email": 26,
        "Verification_Status": 16,
        "Review_Notes": 80,
        "Checked_Date": 12,
    }

    ws = wb.active
    ws.title = "Verified_Firms"
    fill_sheet(ws, rows, widths)

    ws2 = wb.create_sheet("US_IP_Shortlist")
    fill_sheet(ws2, shortlist, widths)

    ws3 = wb.create_sheet("Summary")
    ws3.column_dimensions["A"].width = 42
    ws3.column_dimensions["B"].width = 110
    wrap = Alignment(wrap_text=True, vertical="top")
    summary_rows = [
        ("Checked date", meta.get("checked_date", "")),
        ("Input source", meta.get("input_source", "")),
        ("Method", meta.get("method", "")),
        ("Email policy", meta.get("email_policy", "")),
    ] + [(k, str(v)) for k, v in counts.items()]
    for r_idx, (key, value) in enumerate(summary_rows, start=1):
        a = ws3.cell(row=r_idx, column=1, value=key)
        a.font = Font(bold=True)
        a.alignment = wrap
        b = ws3.cell(row=r_idx, column=2, value=value)
        b.alignment = wrap

    wb.save(OUT_XLSX)
    return True


def summarise(rows: list[dict[str, str]]) -> dict[str, int]:
    ip = Counter(r["IP_Status"] for r in rows)
    us = Counter(r["US_Based_Status"] for r in rows)
    web = Counter(r["Website_Status"] for r in rows)
    ver = Counter(r["Verification_Status"] for r in rows)
    shortlist = sum(
        1 for r in rows if r["US_Based_Status"] == "US-Based" and r["IP_Status"] == "IP"
    )
    shortlist_verified = sum(
        1
        for r in rows
        if r["US_Based_Status"] == "US-Based"
        and r["IP_Status"] == "IP"
        and r["Verification_Status"] == "Verified"
    )
    result: dict[str, int] = {
        "Rows processed": len(rows),
        "IP_Status = IP": ip.get("IP", 0),
        "IP_Status = NIP": ip.get("NIP", 0),
        "IP_Status = Needs Review": ip.get("Needs Review", 0),
        "US_Based_Status = US-Based": us.get("US-Based", 0),
        "US_Based_Status = Non-US": us.get("Non-US", 0),
        "US_Based_Status = Needs Review": us.get("Needs Review", 0),
        "US-Based + IP (shortlist sheet)": shortlist,
        "US-Based + IP + Verified": shortlist_verified,
        "Owner_Email verified & filled": sum(1 for r in rows if r["Owner_Email"]),
        "General_Email filled": sum(1 for r in rows if r["General_Email"]),
        "Verification_Status = Verified": ver.get("Verified", 0),
        "Verification_Status = Needs Review": ver.get("Needs Review", 0),
    }
    for status, n in sorted(web.items()):
        result[f"Website_Status = {status}"] = n
    return result


def parse_args() -> None:
    global INPUT_CSV, RESULTS_JSON, OUT_DIR, OUT_CSV, OUT_XLSX
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--input", type=Path, help="input CSV (default: input_domains.csv)")
    parser.add_argument("--results", type=Path, help="results JSON (default: results.json)")
    parser.add_argument("--out-dir", type=Path, help="output directory (default: output/)")
    parser.add_argument(
        "--prefix", help="output file prefix (default: verified_ip_firms)"
    )
    args = parser.parse_args()
    if args.input:
        INPUT_CSV = args.input if args.input.is_absolute() else HERE / args.input
    if args.results:
        RESULTS_JSON = args.results if args.results.is_absolute() else HERE / args.results
    out_dir = OUT_DIR
    if args.out_dir:
        out_dir = args.out_dir if args.out_dir.is_absolute() else HERE / args.out_dir
    prefix = args.prefix or "verified_ip_firms"
    OUT_DIR = out_dir
    OUT_CSV = out_dir / f"{prefix}.csv"
    OUT_XLSX = out_dir / f"{prefix}.xlsx"


def main() -> int:
    parse_args()
    inputs, results = load_inputs()
    rows = merge_rows(inputs, results)
    validate(rows)
    write_csv(rows)
    xlsx_ok = write_xlsx(rows, results.get("meta", {}))
    print(f"Wrote {OUT_CSV} ({len(rows)} rows)")
    if xlsx_ok:
        print(f"Wrote {OUT_XLSX} (sheets: Verified_Firms, US_IP_Shortlist, Summary)")
    print("\nSummary counts:")
    for key, value in summarise(rows).items():
        print(f"  {key}: {value}")
    print("\nOwner emails (verified on official site only):")
    for row in rows:
        if row["Owner_Email"]:
            print(
                f"  Row {row['Row_ID']:>2} {row['Domain']:<34} "
                f"{row['Owner_First_Name']} {row['Owner_Last_Name']} <{row['Owner_Email']}>"
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
