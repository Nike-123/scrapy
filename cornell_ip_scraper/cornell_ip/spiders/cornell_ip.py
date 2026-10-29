"""Spider: U.S. Intellectual Property law firms from the Cornell LII
Attorney Directory (https://lawyers.law.cornell.edu/lawyers/intellectual-property)

Crawl model
-----------
1. One listing page per U.S. state (+ Washington, D.C.) under
   /lawyers/intellectual-property/<state-slug>.  Each listing page contains
   claimed IP-lawyer profiles for that state (sponsored + regular cards).
2. Every distinct lawyer profile (/lawyer/<slug>) is visited once.
3. A *firm* record is only emitted when the profile identifies a law firm:
   the firm name printed in the profile's "Contact & Map" block (falling
   back to the current position in "Professional Experience").  Solo
   practitioners with no identifiable firm are skipped, per the task spec.
4. City / state come from the profile's Contact & Map address (the
   lawyer's primary office on the directory), falling back to the
   listing card's "<City>, <ST> Intellectual Property Attorney" line.
5. The official website is the directory's "View Website" link for that
   lawyer/firm (listing card and/or profile "Websites & Blogs" section).
   Nothing is guessed: if no website is shown, the field stays blank.

Duplicates (same firm via several lawyers / offices / states) are removed
by cornell_ip.pipelines.DedupPipeline; the Excel file is written by
cornell_ip.pipelines.ExcelWriterPipeline (or build_excel.py offline).
"""
import re

import scrapy

from cornell_ip.items import FirmItem

BASE = "https://lawyers.law.cornell.edu"

# slug -> (State name as printed in the Excel file, state abbreviation)
STATES = {
    "alabama": ("Alabama", "AL"), "alaska": ("Alaska", "AK"),
    "arizona": ("Arizona", "AZ"), "arkansas": ("Arkansas", "AR"),
    "california": ("California", "CA"), "colorado": ("Colorado", "CO"),
    "connecticut": ("Connecticut", "CT"), "delaware": ("Delaware", "DE"),
    "florida": ("Florida", "FL"), "georgia": ("Georgia", "GA"),
    "hawaii": ("Hawaii", "HI"), "idaho": ("Idaho", "ID"),
    "illinois": ("Illinois", "IL"), "indiana": ("Indiana", "IN"),
    "iowa": ("Iowa", "IA"), "kansas": ("Kansas", "KS"),
    "kentucky": ("Kentucky", "KY"), "louisiana": ("Louisiana", "LA"),
    "maine": ("Maine", "ME"), "maryland": ("Maryland", "MD"),
    "massachusetts": ("Massachusetts", "MA"), "michigan": ("Michigan", "MI"),
    "minnesota": ("Minnesota", "MN"), "mississippi": ("Mississippi", "MS"),
    "missouri": ("Missouri", "MO"), "montana": ("Montana", "MT"),
    "nebraska": ("Nebraska", "NE"), "nevada": ("Nevada", "NV"),
    "new-hampshire": ("New Hampshire", "NH"), "new-jersey": ("New Jersey", "NJ"),
    "new-mexico": ("New Mexico", "NM"), "new-york": ("New York", "NY"),
    "north-carolina": ("North Carolina", "NC"),
    "north-dakota": ("North Dakota", "ND"), "ohio": ("Ohio", "OH"),
    "oklahoma": ("Oklahoma", "OK"), "oregon": ("Oregon", "OR"),
    "pennsylvania": ("Pennsylvania", "PA"),
    "rhode-island": ("Rhode Island", "RI"),
    "south-carolina": ("South Carolina", "SC"),
    "south-dakota": ("South Dakota", "SD"), "tennessee": ("Tennessee", "TN"),
    "texas": ("Texas", "TX"), "utah": ("Utah", "UT"),
    "vermont": ("Vermont", "VT"), "Virginia": ("Virginia", "VA"),
    "washington": ("Washington", "WA"),
    "west-virginia": ("West Virginia", "WV"), "wisconsin": ("Wisconsin", "WI"),
    "wyoming": ("Wyoming", "WY"),
    # Washington, D.C. is listed under /dc/washington-dc on the directory.
    "dc": ("District of Columbia", "DC"),
}

US_STATES_BY_ABBR = {abbr: name for name, abbr in STATES.values()}

_CITY_STATE_RE = re.compile(
    r"([A-Z][A-Za-z .'-]+?),\s*([A-Z]{2})\s+Intellectual Property Attorney")
_ZIP_RE = re.compile(r"^(.+?),\s*([A-Z]{2})\s*\d{5}(?:-\d{4})?$")
_PROFILE_RE = re.compile(r"/lawyer/[a-z0-9\-]+$")


def _clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


class CornellIpSpider(scrapy.Spider):
    name = "cornell_ip"
    allowed_domains = ["lawyers.law.cornell.edu"]
    custom_settings = {
        "ITEM_PIPELINES": {
            "cornell_ip.pipelines.DedupPipeline": 100,
            "cornell_ip.pipelines.ExcelWriterPipeline": 200,
        }
    }

    def start_requests(self):
        for slug in STATES:
            if slug == "dc":
                url = f"{BASE}/lawyers/intellectual-property/dc/washington-dc"
            else:
                url = f"{BASE}/lawyers/intellectual-property/{slug}"
            yield scrapy.Request(url, callback=self.parse_listing,
                                 meta={"state_slug": slug})

    # -- listing pages ----------------------------------------------------
    def parse_listing(self, response):
        slug = response.meta["state_slug"]
        state_name, _abbr = STATES[slug]
        seen_profiles = set()
        # Every lawyer-profile link on the page (sponsored + regular cards).
        for href in response.css("a::attr(href)").getall():
            path = href.split("#")[0].split("?")[0].rstrip("/")
            if _PROFILE_RE.search(path) and path not in seen_profiles:
                seen_profiles.add(path)
                yield response.follow(
                    path, callback=self.parse_profile,
                    meta={"state": state_name, "listing_url": response.url,
                          "website_hint": self._card_website(response, path),
                          "city_hint": self._card_city(response, path)})
        # Follow city sub-pages of this state once (they re-list the same
        # lawyers in smaller sets; Scrapy's dupefilter + the dedup pipeline
        # keep the output unique, first occurrence wins).
        for href in response.css("a::attr(href)").getall():
            if re.search(rf"/intellectual-property/{re.escape(slug)}/[a-z\-]+$",
                         href.split("#")[0]):
                yield response.follow(href, callback=self.parse_listing,
                                      meta={"state_slug": slug})

    @staticmethod
    def _card_website(response, profile_path):
        """'View Website' href inside the card that links to profile_path."""
        for card in response.xpath(
                f'//a[contains(@href,"{profile_path}")]/ancestor::*[self::div or self::li][1]'):
            site = card.xpath('.//a[text()[contains(.,"View Website")]]/@href').get()
            if site:
                return site
        return ""

    @staticmethod
    def _card_city(response, profile_path):
        for card in response.xpath(
                f'//a[contains(@href,"{profile_path}")]/ancestor::*[self::div or self::li][1]'):
            m = _CITY_STATE_RE.search(_clean(" ".join(card.xpath(".//text()").getall())))
            if m:
                return _clean(m.group(1)).title()
        return ""

    # -- lawyer profile pages ----------------------------------------------
    def parse_profile(self, response):
        lines = [_clean(t) for t in response.xpath("//body//text()").getall()]
        lines = [ln for ln in lines if ln]

        firm, city, state = self._contact_block(lines)
        if not firm:  # fallback: current position, e.g. "Partner <Firm> LLP"
            firm = self._experience_firm(response)
        if not firm:
            self.logger.info("No identifiable firm, skipping solo profile: %s",
                             response.url)
            return

        if not city:  # fallback to the listing-card city / state of crawl
            city = response.meta.get("city_hint", "")
            state = state or response.meta.get("state", "")

        website = self._profile_website(response) or response.meta.get("website_hint", "")
        yield FirmItem(
            firm_name=firm,
            website=website or "",
            state=state or response.meta.get("state", ""),
            city=city or "",
            lawyer_name=_clean(response.css("h1::text").get()),
            profile_url=response.url,
            listing_url=response.meta.get("listing_url", ""),
        )

    @staticmethod
    def _contact_block(lines):
        """Parse the 'Contact & Map' block:

            Contact & Map
            <Firm Name>            <- optional; absent for solo lawyers
            1201 N. Market St.
            Wilmington, DE 19801
        Returns (firm, city, state_name)."""
        try:
            i = lines.index("Contact & Map")
        except ValueError:
            return "", "", ""
        block = lines[i + 1:i + 10]
        for j, ln in enumerate(block):
            m = _ZIP_RE.match(ln)
            if m:
                city = _clean(m.group(1)).title()
                state = US_STATES_BY_ABBR.get(m.group(2), m.group(2))
                # Firm = first line of the block unless that line is itself
                # a street address (contains a digit) -> solo practitioner.
                firm = ""
                if j >= 2 and not re.search(r"\d", block[0]):
                    firm = block[0]
                elif j >= 2 and re.search(r"\d", block[0]):
                    firm = ""
                return firm, city, state
        return "", "", ""

    @staticmethod
    def _experience_firm(response):
        text = _clean(" ".join(response.xpath(
            '//*[contains(text(),"Professional Experience")]/following::text()').getall()))
        m = re.search(r"(?:Partner|Shareholder|Member|Principal|Counsel|Associate|"
                      r"Attorney|Owner|Founder)\s+([A-Z][\w&.,'’\- ]+?(?:LLP|LLC|P\.C\.|P\.A\.|PLLC|Inc\.|Ltd\.))",
                      text)
        return _clean(m.group(1)) if m else ""

    @staticmethod
    def _profile_website(response):
        site = response.xpath(
            '//a[text()[contains(.,"Website")]]/@href').get()
        return site or ""
