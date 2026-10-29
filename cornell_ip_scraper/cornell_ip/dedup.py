"""Deduplication helpers (pure Python, shared by the Scrapy pipeline,
the offline Excel builder and the tests).

Rules (per task spec):
* Duplicate = same *normalized* firm name OR same official-website domain.
* Normalization is exact after cleaning -- firms are NOT merged merely
  because their names are similar.
* Keep the first occurrence; preserve original collection order.
"""
import re
from urllib.parse import urlparse

_SUFFIXES = {
    "llp", "llc", "llp.", "pllc", "pc", "p.c.", "pa", "p.a.", "inc", "inc.",
    "ltd", "ltd.", "esq", "esqs", "corp", "corp.", "co", "co.", "company",
    "law", "offices", "office", "group", "firm",
}
# NOTE: trailing generic words are only stripped when they follow a comma or
# are legal-entity suffixes; handled conservatively below.


def normalize_firm_name(name: str) -> str:
    if not name:
        return ""
    n = name.lower().strip()
    n = re.sub(r"\bp\.\s*c\.", " pc ", n)
    n = re.sub(r"\bp\.\s*a\.", " pa ", n)
    n = n.replace("&", " and ")
    n = re.sub(r"[.,'’\-/()]", " ", n)
    tokens = [t for t in n.split() if t]
    # Strip trailing legal-entity suffixes only (llp/llc/pc/pa/pllc/inc/ltd/esq)
    entity = {"llp", "llc", "pllc", "pc", "pa", "inc", "ltd", "esq", "esqs", "corp"}
    while tokens and tokens[-1] in entity:
        tokens.pop()
    return " ".join(tokens)


def website_domain(url: str) -> str:
    if not url:
        return ""
    u = url.strip()
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    host = (urlparse(u).hostname or "").lower()
    return host[4:] if host.startswith("www.") else host


def deduplicate(records):
    """records: iterable of dicts with firm_name / website keys.
    Returns (unique_records, duplicate_records)."""
    seen_names, seen_domains = {}, {}
    unique, dupes = [], []
    for rec in records:
        name_key = normalize_firm_name(rec.get("firm_name", ""))
        domain_key = website_domain(rec.get("website", ""))
        dupe_of = None
        if name_key and name_key in seen_names:
            dupe_of = seen_names[name_key]
        elif domain_key and domain_key in seen_domains:
            dupe_of = seen_domains[domain_key]
        if dupe_of is not None:
            dupes.append({**rec, "_duplicate_of_firm": dupe_of.get("firm_name", "")})
            continue
        if name_key:
            seen_names[name_key] = rec
        if domain_key:
            seen_domains[domain_key] = rec
        unique.append(rec)
    return unique, dupes
