# US-Based Law Firm / IP Company Verification — 50 Domains

Verification date: **2026-10-29**. Branch: `arena/da1fcf66-scrapy`.

## Deliverables

| File | Contents |
|---|---|
| `output/verified_ip_firms.xlsx` | **Primary deliverable.** Sheets: `Verified_Firms` (all 50 rows), `US_IP_Shortlist` (US-Based + IP only, 26 rows), `Summary` (counts + method). Frozen header, auto-filter, wrapped notes. |
| `output/verified_ip_firms.csv` | Same 21 columns as the `Verified_Firms` sheet, for import anywhere (including Google Sheets: File → Import). |
| `input_domains.csv` | The 50-row input exactly as received (see Input note below). |
| `results.json` | Machine-readable findings + methodology metadata; source of truth for the build. |
| `build_output.py` | Validates `results.json` against the task rules and regenerates CSV/XLSX. Fails loudly on enum errors or an unsourced/generic `Owner_Email`. |
| `spider_ip_verify.py` | Polite Scrapy spider to re-crawl the domains locally and emit *candidate* evidence for human review (never auto-fills `Owner_Email`). |

## Input note (read first)

The task said "use the existing spreadsheet", but **no spreadsheet was found**:
there is no `.xlsx/.csv` data file in this repository, and the connected
Google Sheets account returned **0 spreadsheets** when listed on 2026-10-29.
The 50 domains concatenated at the top of the task message were therefore
treated as the input list. `Website_URL_Original` preserves each value
verbatim (original casing included); `Company_Name_Original` is blank because
no company names were supplied — `Company_Name_Verified` holds the name taken
from the official website instead. If you share the original spreadsheet (file
or Google Sheets link/ID), these results can be merged back onto your exact
rows/columns by `Domain` without re-verification.

## How each row was verified

1. Live homepage fetch for every domain (final/redirected URL recorded as
   `Verified_Official_Website`; parked, placeholder, retired, broken and
   inaccessible states distinguished in `Website_Status`).
2. Where the homepage did not settle IP services, U.S. presence or ownership,
   the firm's own Practice Areas / Contact / Practitioners subpages were
   fetched. This mattered: Northfield, Schoen and Demeo list IP **only** on
   practice-area pages and would be misclassified from homepages alone.
3. Only where the official site was inaccessible or text-poor were
   professional directories (Martindale-Hubbell, Super Lawyers, Avvo,
   Lawyers.com) consulted. Anything resting on a secondary source is flagged
   `Verification_Status = Needs Review` with the reason in `Review_Notes`.
4. Classification followed the task rules strictly:
   - `IP_Status`: `IP` / `NIP` / `Needs Review`. Inaccessible sites were
     **never** auto-marked `NIP`; `NIP` was used only with positive evidence
     (e.g., a complete official services list with no IP work).
   - `US_Based_Status`: `US-Based` / `Non-US` / `Needs Review`. A `.com` or a
     U.S. phone number alone was never treated as proof; geography is flagged
     separately from IP relevance (one confirmed `Non-US` IP firm: De Beer
     Attorneys, South Africa).
   - Owner: founder/owner/managing partner/principal where the site (or a
     flagged secondary source) identifies one. Where two equal name partners
     exist with no primary (Davidson Sheehan LLP), owner fields were left
     blank and both partners' published emails placed in `Review_Notes` for a
     human decision. First-listed attorneys were never assumed to be owners
     (e.g., Flume Law Firm).
   - `Owner_Email`: filled **only** when published on the official website in
     direct association with the named person (**6** firms). No address was
     ever generated from a naming pattern. Generic mailboxes went to
     `General_Email` (**5** firms). Snippet/directory-only email candidates
     (e.g., tamarofflaw.com) appear only as clearly-labelled unverified leads
     in `Review_Notes`. Referral-attorney and staff emails that appear on two
     sites (armijolawfirm.com, ojchamber.com) were deliberately excluded.

Sandbox note: this environment's direct outbound HTTP is limited to a small
host allowlist (GitHub/npm/PyPI), so `curl`/Scrapy cannot reach firm websites
from here (verified empirically). Evidence above was gathered with the agent
page-fetch/search tools; `spider_ip_verify.py` reproduces the crawl on an
unrestricted network.

## Results summary (2026-10-29)

- Rows processed: **50**
- `IP_Status`: **IP 28** · **NIP 7** · **Needs Review 15**
- `US_Based_Status`: **US-Based 30** · **Non-US 1** · **Needs Review 19**
- **US-Based + IP shortlist: 26** firms, of which **18** are fully
  `Verified` on live official-site evidence (the other 8 carry a specific
  confirmation note, e.g., IP evidenced via Martindale or street address not
  published).
- `Verification_Status`: Verified 22 · Needs Review 28
- Websites: Active 36 · Active-Redirected 1 · Active-Retired 1 ·
  Inaccessible 8 · Broken/Placeholder 2 · Parked 2
- Verified owner emails: **6** — Northfield, Abel, Snyder (Tampa),
  Deutch, Janowski (DevicePatent), Diament. Full addresses + source URLs are
  in the spreadsheet.

Notable data-quality catches: `arthurlaw.com` is listed for sale;
`armijolawfirm.com` retired in 2019; `demfirm.com` shows signs of
template/SEO compromise (flagged, unverified); `dkpls.com` appears to be a
duplicate/legacy domain of the Kleiman Law firm already captured via
`selki.com` → `kleiman.law` (counted once); `deenakalai.com` is verified IP
but on sabbatical through 2026; `devicepatent.com` and
`northwestpatent.com` provide patent services without clearly stated
attorney/agent licensure parity (Northwest is a registered patent *agent*;
DevicePatent's licensure is unstated — both flagged).

## Re-run / refresh

```bash
# Regenerate deliverables after editing results.json
python3 -m venv /tmp/venv && /tmp/venv/bin/pip install openpyxl
/tmp/venv/bin/python ip_verification/build_output.py

# Re-crawl candidates locally (unrestricted network; respects robots.txt)
pip install scrapy   # or use this checkout: pip install -e .
cd ip_verification
scrapy runspider spider_ip_verify.py -O output/candidates.jl
```

Freshness: law-firm sites, sabbaticals, retirements and domain sales change
over time. Re-verify `Needs Review` rows (especially the 8 inaccessible
domains) from a normal browser before any outreach, and re-check the two
secondary-source IP rows (Daniel Law Office PC, David Steiner & Associates)
on the firms' own practice pages.

## Compliance note

Only publicly published professional information was recorded. Respect each
site's terms/robots rules and applicable email/outreach law (e.g., CAN-SPAM)
before contacting anyone in the shortlist; inclusion here is a research
finding, not consent to marketing.

## Batch 2 (Rows 51–110, verified 2026-10-30)

A second list of 60 `Company Name + Website` pairs was supplied and verified with the
identical rules. Inputs/findings/outputs:

```
batch2/input_batch2.csv
batch2/results_batch2.json                      <- per-row findings (source of truth)
batch2/output/verified_ip_firms_batch2.xlsx     (Verified_Firms, US_IP_Shortlist, Summary)
batch2/output/verified_ip_firms_batch2.csv
input_all.csv / results_all.json                <- combined Rows 1-110 build inputs
output/verified_ip_firms_all.xlsx / .csv        <- combined master (Rows 1-110)
```

Company names in the paste arrived without spaces (e.g. `Dvorak&Donovan`); they were
space-normalized in `input_batch2.csv` while the URLs are preserved verbatim.
`build_output.py` now takes `--input/--results/--out-dir/--prefix` (defaults =
Batch 1).

### Batch 2 summary (60 rows)

| Classification | Count |
|---|---|
| IP | 31 |
| NIP (verified, no IP service) | 13 |
| Needs Review | 16 |
| US-Based | 29 |
| Non-US | 1 (elizabethchurchill.com — Abu Dhabi academic, not a lawyer) |
| U.S. Needs Review | 30 |
| **US-Based + IP shortlist** | **22** (15 fully Verified) |
| Owner emails verified on official site | 5 |
| General emails | 11 |

Combined Rows 1–110: IP 59 / NIP 20 / Needs Review 31; US-Based 59; US + IP
shortlist 48; verified owner emails 11.

### Batch 2 verified owner emails (official site only)

| Row | Person | Email | Firm |
|---|---|---|---|
| 79 | Eva Guo | evaguo8@gmail.com (Gmail, NIP firm) | evaguo.com (personal injury) |
| 86 | Michael Fabiano | mfabiano@FabianoLawFirm.com | fabianolawfirm.com |
| 87 | Christopher Falkowski | Chris@FalkowskiPLLC.com | falkowskipllc.com |
| 101 | Paul Fattibene | paul@fattibene.com | fattibene.com |
| 109 | Jeffrey Finn | jeff@finniplaw.com | finniplaw.com |

### Batch 2 notable findings / input-list errors

- **Repurposed/hijacked domains (do not use):** `uebleriplaw.com` → Indonesian
  gambling site; `faylaw.com` → door-mat e-commerce spam; `file-ee-patents.com` →
  a Calgary Pilates studio. `extrudedelement.com` is an unconfigured-host page.
- **Wrong URL for the named company:** `ferguson.com` (plumbing supplies, 404) vs
  *Fergason* Patent Properties; `fipp.com` (global media association);
  `lawcrossing.com` (jobs board) vs Ecoff Landsberg; `rosenlaw.com` (Lawrence
  Rosen, retired) vs Michael B. Einschlag; `echelaw.com` ("Eche Law" blog) vs
  Echelon Law Group; `elizabethchurchill.com` (HCI professor, UAE);
  `docdreyfus.com` appears tied to a PhD psychologist listing, not "Edward
  Dreyfus, Esq." (two different people — do not merge).
- **Input name vs. actual practice:** "Elizabeth Kim Patent Law Offices"
  (kimassociates.com) — official practice list has no patent/IP work (NIP);
  Ferrance Law is now NM civil appeals ("not taking new cases") despite the
  founder's USPTO background; "Dvorak & Donovan" is now Dvorak Law Group.
- **Merged/redirected:** Fathom Law → merged into Scale LLP; Esposito →
  elglawfirm.com (NIP); Evans Fears → efsmmlaw.com (NIP).
- **Not law firms:** DW Gildart (engineering), Epicenter IP (patent brokerage),
  Fine Point Patent (drawings), FastPatentPartner (legal-tech product), FirmForward
  (law-firm marketing). Epsilon Patents / Express Search are patent-search firms
  (counted IP per the task's search-services list, flagged as non-firms).
- **Inaccessible (Needs Review, never auto-NIP):** dykaslaw.com, docdreyfus.com,
  wengreenpatentlaw.com (500; attorney may now be at Knobbe Martens), faricylaw.com
  (500), focusonpatents.com. Directory leads are in those rows' Review_Notes.
