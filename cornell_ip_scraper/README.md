# Cornell LII — U.S. Intellectual Property Law Firms Scraper

Collects **unique law firms** (verified firm name + official website URL)
from the Cornell Law School / Justia Attorney Directory, following the
hierarchy:

```
Main Directory -> State -> City -> Directory Pages -> Law Firms
```

Starting URL: https://lawyers.law.cornell.edu/lawyers/intellectual-property

## Rules enforced

- Final workbook contains **exactly two columns** — `Law Firm Name` and
  `Official Website URL`. Lawyer names, phones, addresses, state and city
  are kept only in internal tracking files for verification/auditing.
- A record is emitted only when a profile identifies an actual law firm
  (Contact & Map firm heading, firm-named website label, or a current
  firm position). Solo practitioners, in-house lawyers and conflicting /
  unverifiable firm associations are excluded (and logged in the raw
  store with an explanatory note).
- Websites are taken verbatim from the directory ("View Website" /
  profile "Websites & Blogs") or the firm's own site — never invented;
  unverifiable websites stay blank.
- Sponsored "Serving <State/City>" text is **never** used as location
  proof; state/city come only from a real "<City>, <ST>" listing card or
  the profile's Contact & Map address.
- Deduplication is **global** (across all states/cities) by normalized
  firm name OR official website domain; first occurrence wins; merely
  similar firm names are never merged.

## Run the crawl (normal network)

```bash
cd cornell_ip_scraper
scrapy crawl cornell_ip -s JOBDIR=data/crawljob   # resumable
```

`ROBOTSTXT_OBEY = True`, a download delay and AutoThrottle are set in
`cornell_ip/settings.py`. With `JOBDIR` set, interrupting and re-running
the same command resumes the scheduler queue instead of starting over;
every scraped record is appended immediately to `data/raw_records.jsonl`
by `JsonlRawStorePipeline`, so no collected data is lost.

> Note: direct HTTPS to `lawyers.law.cornell.edu` is blocked inside the
> sandbox used to build this project, so the current dataset was
> harvested through a page-fetch channel into the same raw store. The
> spider implements the identical extraction logic for open networks.

## Rebuild the final workbook (offline, any time)

```bash
python build_workbook.py
```

This merges the legacy `data/raw_records.json` with
`data/raw_records.jsonl` (union by profile URL — existing data is never
overwritten), deduplicates globally and writes:

| File | Contents |
|---|---|
| `output/IP_Law_Firms.xlsx` | Final workbook: one sheet `IP Law Firms`, exactly `Law Firm Name` + `Official Website URL`, frozen header, filters on, clickable verified URLs, blanks left blank |
| `data/raw_records.jsonl` | All accumulated raw records (incl. excluded profiles, with reasons in `note`) |
| `data/state_progress.json` | Per-state status (`not_started` / `in_progress` / `completed` / `failed_or_inaccessible` / `no_relevant_listings_found`), pages processed, cities seen |
| `data/city_progress.json` | Per-(state, city) processing status and record counts |
| `data/errors.jsonl` | Failed / inaccessible pages with reasons |
| `data/collection_summary.json` | Quality-check counts for the latest build |

`build_excel.py` remains as the legacy builder for the earlier detailed
5-column audit workbook (`US_IP_Law_Firms_Cornell.xlsx` at repo root).

## Tests

```bash
PYTHONPATH=. python -m pytest tests/ -q
```

## Known coverage limitations

- **North Dakota:** the directory listing itself returns "no results" —
  no IP lawyer listings exist to process (`no_relevant_listings_found`).
- **Alaska / Wyoming:** listings contain only solo or in-house lawyers
  with no identifiable firm; excluded under the rules above.
- Pagination on the largest states (CA, NY, FL, TX, …) is deep and
  sponsored listings repeat across pages; states are marked `completed`
  in `state_progress.json` only once their pagination has genuinely been
  exhausted — until then they honestly remain `in_progress`. No complete
  U.S. coverage is claimed.
