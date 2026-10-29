#!/usr/bin/env python3
"""Polite Scrapy spider that re-checks the input domains and emits *candidate*
evidence for human verification.

Why "candidate": IP-service and U.S.-base classification and, above all,
owner-email attribution require human judgement (see README.md). This spider
therefore never writes Owner_Email; it records every email it finds together
with the page URL it was found on, so a reviewer can promote only addresses
that the official site explicitly associates with the named owner. It never
synthesises an address from a naming pattern.

Politeness / compliance:
    * ROBOTSTXT_OBEY is on; AutoThrottle is on; one concurrent request per
      domain; a descriptive User-Agent; at most MAX_PAGES_PER_SITE pages per
      input domain (homepage + linked About/Contact/Practice/Services/Team
      pages on the same registered domain, following redirects).
    * Run it where outbound web access is permitted (the Arena sandbox only
      allows a small host allowlist, which is why the verification in
      results.json was gathered with the agent fetch/search tools instead).

Usage (from the repository root, using this checkout's Scrapy):
    pip install -e .            # or: pip install scrapy
    cd ip_verification
    scrapy runspider spider_ip_verify.py -O output/candidates.jl

Input:  input_domains.csv next to this file (Domain column).
Output: one JSON-lines record per input domain with final URL, status,
        page titles, US/IP keyword evidence snippets and found emails.
"""

from __future__ import annotations

import csv
import re
from pathlib import Path
from urllib.parse import urlparse

import scrapy

HERE = Path(__file__).resolve().parent
INPUT_CSV = HERE / "input_domains.csv"
MAX_PAGES_PER_SITE = 6

IP_KEYWORDS = [
    "intellectual property",
    "patent",
    "trademark",
    "copyright",
    "trade secret",
    "freedom to operate",
    "freedom-to-operate",
    "patentability",
    "portfolio management",
    "ip licensing",
    "patent prosecution",
]
US_KEYWORDS = [
    "united states", ", al ", ", ak ", ", az ", ", ar ", ", ca ", ", co ",
    ", ct ", ", de ", ", fl ", ", ga ", ", hi ", ", id ", ", il ", ", in ",
    ", ia ", ", ks ", ", ky ", ", la ", ", me ", ", md ", ", ma ", ", mi ",
    ", mn ", ", ms ", ", mo ", ", mt ", ", ne ", ", nv ", ", nh ", ", nj ",
    ", nm ", ", new York", ", nc ", ", nd ", ", oh ", ", ok ", ", or ",
    ", pa ", ", ri ", ", sc ", ", sd ", ", tn ", ", tx ", ", ut ", ", vt ",
    ", va ", ", wa ", ", wi ", ", wv ", ", wy ", ", dc ",
]
FOLLOW_LINK_HINTS = (
    "about", "contact", "practice", "service", "attorney", "team",
    "people", "practitioner", "profile", "firm", "expertise",
)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
GENERIC_LOCALPARTS = {
    "info", "contact", "mailbox", "office", "admin", "hello", "support",
    "sales", "inquiries", "inquiry",
}


def _snippets(text: str, keywords: list[str], *, limit: int = 5) -> list[str]:
    low = text.lower()
    found: list[str] = []
    for kw in keywords:
        start = 0
        while len(found) < limit:
            idx = low.find(kw.lower(), start)
            if idx == -1:
                break
            snippet = " ".join(text[max(0, idx - 80): idx + 120].split())
            if snippet not in found:
                found.append(snippet)
            start = idx + len(kw)
        if len(found) >= limit:
            break
    return found


class IpVerifySpider(scrapy.Spider):
    name = "ip_verify"
    custom_settings = {
        "ROBOTSTXT_OBEY": True,
        "AUTOTHROTTLE_ENABLED": True,
        "AUTOTHROTTLE_START_DELAY": 1.0,
        "AUTOTHROTTLE_MAX_DELAY": 10.0,
        "CONCURRENT_REQUESTS_PER_DOMAIN": 1,
        "DOWNLOAD_DELAY": 1.0,
        "USER_AGENT": (
            "IPFirmVerificationBot/1.0 (+manual research; "
            "respects robots.txt; contact via repository owner)"
        ),
        "LOG_LEVEL": "INFO",
    }

    def start_requests(self):
        with INPUT_CSV.open(newline="", encoding="utf-8") as fh:
            for row in csv.DictReader(fh):
                domain = (row.get("Domain") or "").strip().lower()
                if not domain:
                    continue
                item: dict = {
                    "Row_ID": row.get("Row_ID", ""),
                    "Domain": domain,
                    "pages": [],
                    "emails_found": [],
                    "ip_evidence_snippets": [],
                    "us_evidence_snippets": [],
                    "_pages_seen": 0,
                }
                yield scrapy.Request(
                    f"https://www.{domain}/",
                    callback=self.parse_page,
                    errback=self.on_error,
                    cb_kwargs={"item": item, "is_home": True},
                    dont_filter=True,
                )

    def on_error(self, failure, item=None):  # noqa: ARG002 - scrapy errback API
        if item is None:
            return
        item["error"] = repr(failure.value)
        item["candidate_Website_Status"] = "Inaccessible (see error)"
        yield self._finalise(item)

    def parse_page(self, response, item, is_home=False):
        item["_pages_seen"] += 1
        title = " ".join(response.xpath("//title/text()").getall()).strip()
        text = " ".join(response.xpath("//body//text()").getall())
        text = " ".join(text.split())
        item["pages"].append(
            {"url": response.url, "status": response.status, "title": title}
        )
        if is_home:
            item["candidate_final_url"] = response.url
            item["candidate_Website_Status"] = f"Active (HTTP {response.status})"

        for email in sorted(set(EMAIL_RE.findall(text))):
            local = email.split("@")[0].lower()
            entry = {
                "email": email,
                "found_on": response.url,
                "kind": "generic" if local in GENERIC_LOCALPARTS else "personal-looking",
                "note": "candidate only - human must confirm the named owner association",
            }
            if entry not in item["emails_found"]:
                item["emails_found"].append(entry)

        for snippet in _snippets(text, IP_KEYWORDS):
            if snippet not in item["ip_evidence_snippets"]:
                item["ip_evidence_snippets"].append(snippet)
        for snippet in _snippets(text, US_KEYWORDS):
            if snippet not in item["us_evidence_snippets"]:
                item["us_evidence_snippets"].append(snippet)

        # Follow a small number of same-site About/Contact/Practice links.
        followed = 0
        if item["_pages_seen"] < MAX_PAGES_PER_SITE:
            home_host = urlparse(item.get("candidate_final_url") or response.url).netloc
            for href, label in zip(
                response.xpath("//a/@href").getall(),
                response.xpath("//a//text()").getall(),
                strict=False,
            ):
                if item["_pages_seen"] + followed >= MAX_PAGES_PER_SITE:
                    break
                absolute = response.urljoin(href)
                parsed = urlparse(absolute)
                if parsed.netloc != home_host:
                    continue
                haystack = f"{parsed.path} {label}".lower()
                if not any(hint in haystack for hint in FOLLOW_LINK_HINTS):
                    continue
                if any(p["url"] == absolute for p in item["pages"]):
                    continue
                followed += 1
                yield response.follow(
                    absolute,
                    callback=self.parse_page,
                    cb_kwargs={"item": item, "is_home": False},
                )
        if followed == 0:
            yield self._finalise(item)

    @staticmethod
    def _finalise(item: dict) -> dict:
        item.pop("_pages_seen", None)
        item["candidate_IP_mentions"] = bool(item["ip_evidence_snippets"])
        item["candidate_US_mentions"] = bool(item["us_evidence_snippets"])
        item["human_verification_required"] = True
        item["Owner_Email"] = ""  # never auto-filled; see module docstring
        return item
