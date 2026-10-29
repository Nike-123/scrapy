"""Resumable on-disk storage for the Cornell IP collection.

Layout (relative to the project directory):

    data/raw_records.jsonl    one JSON object per harvested profile record,
                              appended incrementally -- never overwritten
    data/raw_records.json     legacy accumulated file from the first
                              collection rounds (kept for audit; migrated
                              into the JSONL store by build_workbook.py)
    data/state_progress.json  per-state status inventory
    data/city_progress.json   per-(state, city) status inventory
    data/errors.jsonl         failed / inaccessible pages, with reasons

All writers append or merge, so a crawl or build can be interrupted and
resumed without losing previously collected data.
"""
import json
import os
from datetime import datetime, timezone

# Canonical state inventory: slug -> display name.  Washington, D.C. is
# listed by the directory under /dc/washington-dc.
STATES = [
    ("alabama", "Alabama"), ("alaska", "Alaska"), ("arizona", "Arizona"),
    ("arkansas", "Arkansas"), ("california", "California"),
    ("colorado", "Colorado"), ("connecticut", "Connecticut"),
    ("delaware", "Delaware"), ("florida", "Florida"), ("georgia", "Georgia"),
    ("hawaii", "Hawaii"), ("idaho", "Idaho"), ("illinois", "Illinois"),
    ("indiana", "Indiana"), ("iowa", "Iowa"), ("kansas", "Kansas"),
    ("kentucky", "Kentucky"), ("louisiana", "Louisiana"), ("maine", "Maine"),
    ("maryland", "Maryland"), ("massachusetts", "Massachusetts"),
    ("michigan", "Michigan"), ("minnesota", "Minnesota"),
    ("mississippi", "Mississippi"), ("missouri", "Missouri"),
    ("montana", "Montana"), ("nebraska", "Nebraska"), ("nevada", "Nevada"),
    ("new-hampshire", "New Hampshire"), ("new-jersey", "New Jersey"),
    ("new-mexico", "New Mexico"), ("new-york", "New York"),
    ("north-carolina", "North Carolina"),
    ("north-dakota", "North Dakota"), ("ohio", "Ohio"),
    ("oklahoma", "Oklahoma"), ("oregon", "Oregon"),
    ("pennsylvania", "Pennsylvania"), ("rhode-island", "Rhode Island"),
    ("south-carolina", "South Carolina"),
    ("south-dakota", "South Dakota"), ("tennessee", "Tennessee"),
    ("texas", "Texas"), ("utah", "Utah"), ("vermont", "Vermont"),
    ("Virginia", "Virginia"), ("washington", "Washington"),
    ("west-virginia", "West Virginia"), ("wisconsin", "Wisconsin"),
    ("wyoming", "Wyoming"), ("dc", "District of Columbia"),
]
STATE_NAMES = [name for _slug, name in STATES]

STATUS_NOT_STARTED = "not_started"
STATUS_IN_PROGRESS = "in_progress"
STATUS_COMPLETED = "completed"
STATUS_FAILED = "failed_or_inaccessible"
STATUS_NO_LISTINGS = "no_relevant_listings_found"


def _now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def load_jsonl(path):
    records = []
    if os.path.exists(path):
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if line:
                    records.append(json.loads(line))
    return records


def append_jsonl(path, record):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "a", encoding="utf-8") as fh:
        fh.write(json.dumps(record, ensure_ascii=False) + "\n")


def write_jsonl(path, records):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        for rec in records:
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")


def merge_records(legacy_records, jsonl_records):
    """Union of legacy JSON records and JSONL records, preserving every
    raw record for auditing.  Records are keyed by (profile_url,
    listing_url) -- the same lawyer harvested from two different listings
    stays as two raw records (global firm dedup handles it later) --
    with the JSONL version winning on an exact-key conflict.  First-seen
    order is preserved, and re-running the merge is idempotent."""
    merged, order = {}, []
    for rec in list(legacy_records) + list(jsonl_records):
        key = (rec.get("profile_url"), rec.get("listing_url")) \
            if rec.get("profile_url") else json.dumps(rec, sort_keys=True)
        if key not in merged:
            order.append(key)
        merged[key] = rec
    return [merged[k] for k in order]


def log_error(errors_path, url, reason, state="", city=""):
    append_jsonl(errors_path, {
        "timestamp": _now(), "url": url, "state": state, "city": city,
        "reason": reason,
    })


def build_progress(records, exhausted_states=(), no_listing_states=(),
                   failed_states=None):
    """Derive (state_progress, city_progress) from accumulated records.

    A state is 'completed' only when listed in exhausted_states (its
    pagination was followed to the end); states with records but remaining
    pagination stay 'in_progress' -- progress is never overstated.
    """
    failed_states = failed_states or {}
    by_state, cities = {}, {}
    for rec in records:
        state = (rec.get("state") or "").strip()
        if not state:
            continue
        by_state.setdefault(state, []).append(rec)
        city = (rec.get("city") or "").strip()
        if city:
            entry = cities.setdefault((state, city),
                                      {"records": 0, "firms": 0, "pages": set()})
            entry["records"] += 1
            entry["firms"] += 1 if (rec.get("firm_name") or "").strip() else 0
            if rec.get("listing_url"):
                entry["pages"].add(rec["listing_url"])

    state_progress = {}
    for _slug, name in STATES:
        recs = by_state.get(name, [])
        pages = sorted({r["listing_url"] for r in recs if r.get("listing_url")})
        if name in failed_states:
            status = STATUS_FAILED
        elif name in no_listing_states:
            status = STATUS_NO_LISTINGS
        elif name in exhausted_states:
            status = STATUS_COMPLETED
        elif recs:
            status = STATUS_IN_PROGRESS
        else:
            status = STATUS_NOT_STARTED
        state_progress[name] = {
            "status": status,
            "records_collected": len(recs),
            "firm_bearing_records": sum(1 for r in recs
                                        if (r.get("firm_name") or "").strip()),
            "cities_seen": sorted({r["city"] for r in recs if r.get("city")}),
            "listing_pages_processed": pages,
            "listing_pages_processed_count": len(pages),
            "note": failed_states.get(name, ""),
            "updated_at": _now(),
        }

    city_progress = {}
    for (state, city), entry in sorted(cities.items()):
        city_progress[f"{state}::{city}"] = {
            "state": state, "city": city, "status": "processed",
            "records_collected": entry["records"],
            "firm_bearing_records": entry["firms"],
            "listing_pages_processed": sorted(entry["pages"]),
            "updated_at": _now(),
        }
    return state_progress, city_progress


def save_progress(data_dir, state_progress, city_progress):
    os.makedirs(data_dir, exist_ok=True)
    with open(os.path.join(data_dir, "state_progress.json"), "w",
              encoding="utf-8") as fh:
        json.dump(state_progress, fh, indent=2, ensure_ascii=False)
    with open(os.path.join(data_dir, "city_progress.json"), "w",
              encoding="utf-8") as fh:
        json.dump(city_progress, fh, indent=2, ensure_ascii=False)
