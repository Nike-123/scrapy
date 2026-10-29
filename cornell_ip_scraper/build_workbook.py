#!/usr/bin/env python3
"""Canonical offline builder for the final deliverable.

Pipeline (resumable, append-only raw store):

    data/raw_records.json  (legacy accumulated records, preserved)
  + data/raw_records.jsonl (incremental records appended by the spider /
                            harvesting runs)
        -> merge by profile_url (nothing already collected is lost)
        -> data/raw_records.jsonl rewritten with the merged union
        -> global deduplication (normalized firm name / website domain,
           first occurrence wins, similar names never merged)
        -> output/IP_Law_Firms.xlsx   (one sheet 'IP Law Firms', EXACTLY
           two columns: Law Firm Name | Official Website URL)
        -> data/state_progress.json + data/city_progress.json refreshed
        -> data/collection_summary.json refreshed

Direct HTTPS to lawyers.law.cornell.edu is blocked inside this sandbox,
so records were harvested through the page-fetch channel into the raw
store; on a normal network the same store is fed by:

    scrapy crawl cornell_ip -s JOBDIR=data/crawljob

Usage:  python build_workbook.py
Stdout: JSON quality-check summary.
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cornell_ip.dedup import deduplicate            # noqa: E402
from cornell_ip.excel import build_firm_workbook    # noqa: E402
from cornell_ip.store import (                       # noqa: E402
    build_progress, load_jsonl, log_error, merge_records, save_progress,
    write_jsonl,
)

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")

# States whose directory pagination has NOT been followed to the end yet.
# Nothing is marked "completed" until its Next/page links are exhausted;
# build_progress() derives the honest status from this list.
EXHAUSTED_STATES = ()
# Verified against the live directory during collection:
NO_LISTING_STATES = {"North Dakota"}  # listing returns "no results"
FAILED_STATES = {
    "Alaska": "Listings processed; only solo / in-house lawyers found, "
              "no identifiable law firm (excluded per rules)",
    "Wyoming": "Listings processed; only solo / in-house lawyers found, "
               "no identifiable law firm (excluded per rules)",
}


def main():
    legacy_path = os.path.join(DATA, "raw_records.json")
    jsonl_path = os.path.join(DATA, "raw_records.jsonl")
    errors_path = os.path.join(DATA, "errors.jsonl")
    out_path = os.path.join(HERE, "output", "IP_Law_Firms.xlsx")

    legacy = []
    if os.path.exists(legacy_path):
        with open(legacy_path, encoding="utf-8") as fh:
            legacy = json.load(fh)
    records = merge_records(legacy, load_jsonl(jsonl_path))
    write_jsonl(jsonl_path, records)  # merged union; append-only from here

    if not os.path.exists(errors_path):
        log_error(errors_path,
                  "https://lawyers.law.cornell.edu/lawyers/intellectual-property/north-dakota",
                  "Directory listing returns 'Sorry, there were no results "
                  "for this area' -- no IP lawyer listings exist to process",
                  state="North Dakota")

    firms_only = [r for r in records if (r.get("firm_name") or "").strip()]
    excluded = [r for r in records if not (r.get("firm_name") or "").strip()]
    unique, dupes = deduplicate(firms_only)
    build_firm_workbook(unique, out_path)

    state_progress, city_progress = build_progress(
        records, exhausted_states=EXHAUSTED_STATES,
        no_listing_states=NO_LISTING_STATES, failed_states=FAILED_STATES)
    save_progress(DATA, state_progress, city_progress)

    summary = {
        "raw_records_collected": len(records),
        "raw_firm_bearing_records": len(firms_only),
        "excluded_no_identifiable_firm": len(excluded),
        "duplicate_records_removed": len(dupes),
        "unique_firms_remaining": len(unique),
        "firms_with_verified_website": sum(
            1 for r in unique if (r.get("website") or "").strip()),
        "firms_with_blank_website": sum(
            1 for r in unique if not (r.get("website") or "").strip()),
        "states_discovered": len(state_progress),
        "states_with_unique_firms": len(
            {r["state"] for r in unique if r.get("state")}),
        "states_completed": sorted(
            k for k, v in state_progress.items() if v["status"] == "completed"),
        "states_in_progress": sorted(
            k for k, v in state_progress.items() if v["status"] == "in_progress"),
        "states_no_relevant_listings": sorted(
            k for k, v in state_progress.items()
            if v["status"] == "no_relevant_listings_found"),
        "states_failed_or_inaccessible": sorted(
            k for k, v in state_progress.items()
            if v["status"] == "failed_or_inaccessible"),
        "cities_processed": len(city_progress),
        "output": out_path,
        "duplicates": dupes,
    }
    print(json.dumps(summary, indent=2))
    with open(os.path.join(DATA, "collection_summary.json"), "w",
              encoding="utf-8") as fh:
        json.dump(summary, fh, indent=2)


if __name__ == "__main__":
    main()
