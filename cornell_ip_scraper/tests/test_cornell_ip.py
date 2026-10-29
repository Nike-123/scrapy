import json, os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, "/tmp/venv/lib/python3.11/site-packages")
import pytest
from cornell_ip.dedup import deduplicate, normalize_firm_name, website_domain

FIX = os.path.join(os.path.dirname(__file__), "fixtures")

def test_normalize():
    assert normalize_firm_name("Morris, Nichols, Arsht & Tunnell LLP") == "morris nichols arsht and tunnell"
    assert normalize_firm_name("Smith & Jones, P.C.") == "smith and jones"
    assert website_domain("http://www.SVLaw.com/attorneys/x/") == "svlaw.com"
    assert website_domain("") == ""

def test_dedup_first_wins_no_similar_merge():
    recs = [
        {"firm_name": "Alpha Law LLP", "website": "https://alpha.com/", "state": "Delaware", "city": "Wilmington"},
        {"firm_name": "Alpha Law LLC", "website": "https://www.alpha.com/team", "state": "Delaware", "city": "Dover"},
        {"firm_name": "Alpha & Beta LLP", "website": "https://alphabeta.com/", "state": "Delaware", "city": "Newark"},
        {"firm_name": "Beta Law", "website": "https://alpha.com/other", "state": "Delaware", "city": "Dover"},
    ]
    unique, dupes = deduplicate(recs)
    assert [u["firm_name"] for u in unique] == ["Alpha Law LLP", "Alpha & Beta LLP"]
    assert len(dupes) == 2

def test_contact_block_firm_and_solo():
    pytest.importorskip("scrapy")
    from scrapy.http import HtmlResponse, Request
    from cornell_ip.spiders.cornell_ip import CornellIpSpider
    spider = CornellIpSpider()
    for fname, expect_firm in [("profile_firm.html", "Morris, Nichols, Arsht & Tunnell LLP"),
                               ("profile_solo.html", "")]:
        body = open(os.path.join(FIX, fname), "rb").read()
        resp = HtmlResponse(url="https://lawyers.law.cornell.edu/lawyer/x-1",
                            body=body, encoding="utf-8",
                            request=Request(url="https://lawyers.law.cornell.edu/lawyer/x-1",
                                            meta={"state": "Delaware", "city_hint": "", "website_hint": "", "listing_url": ""}))
        items = list(spider.parse_profile(resp))
        if expect_firm:
            assert len(items) == 1
            assert items[0]["firm_name"] == expect_firm
            assert items[0]["city"] == "Wilmington" and items[0]["state"] == "Delaware"
            assert items[0]["website"].startswith("https://www.mnat.com/")
        else:
            assert items == []
