#!/usr/bin/env python3
"""Offline builder: data/raw_records.json -> dedup -> US_IP_Law_Firms_Cornell.xlsx

Used when the crawl records were collected through an alternative channel
(e.g. the sandbox page-fetch tool, because direct HTTPS to
lawyers.law.cornell.edu is blocked in this environment).  The dedup and
Excel code are exactly the ones used by the Scrapy pipelines.

Usage:  python build_excel.py [raw.json] [out.xlsx]
Stdout: JSON summary {raw, duplicates_removed, unique, states_covered, ...}
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cornell_ip.dedup import deduplicate          # noqa: E402
from cornell_ip.excel import build_workbook       # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))


def main():
    raw_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "data", "raw_records.json")
    out_path = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "..", "US_IP_Law_Firms_Cornell.xlsx")
    out_path = os.path.normpath(out_path)
    with open(raw_path, encoding="utf-8") as fh:
        raw = json.load(fh)
    firms_only = [r for r in raw if (r.get("firm_name") or "").strip()]
    unique, dupes = deduplicate(firms_only)
    build_workbook(unique, out_path)
    summary = {
        "raw_records_collected": len(raw),
        "raw_firm_bearing_records": len(firms_only),
        "duplicate_records_removed": len(dupes),
        "unique_firms_remaining": len(unique),
        "states_covered": len({r["state"] for r in unique if r.get("state")}),
        "states_list": sorted({r["state"] for r in unique if r.get("state")}),
        "firms_with_website": sum(1 for r in unique if (r.get("website") or "").strip()),
        "firms_without_website_blank_cell": sum(1 for r in unique if not (r.get("website") or "").strip()),
        "output": out_path,
        "duplicates": dupes,
    }
    print(json.dumps(summary, indent=2))
    with open(os.path.join(HERE, "data", "collection_summary.json"), "w", encoding="utf-8") as fh:
        json.dump(summary, fh, indent=2)


if __name__ == "__main__":
    main()
