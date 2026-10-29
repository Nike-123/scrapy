"""Pipelines: incremental JSONL raw store (resumable), global
deduplication (first occurrence wins) + final two-column Excel writer."""
import os

from cornell_ip.dedup import deduplicate
from cornell_ip.excel import build_firm_workbook
from cornell_ip.store import append_jsonl


def _data_dir(spider):
    return spider.settings.get("CORNELL_DATA_DIR", os.path.join(os.getcwd(), "data"))


class JsonlRawStorePipeline:
    """Append every harvested record to data/raw_records.jsonl as it is
    scraped, so an interrupted crawl loses nothing and can resume."""

    def process_item(self, item, spider):
        append_jsonl(os.path.join(_data_dir(spider), "raw_records.jsonl"),
                     dict(item))
        return item


class DedupPipeline:
    """Collect items, drop exact normalized-name / domain duplicates."""

    def __init__(self):
        self.records = []
        self.stats = {"raw": 0, "duplicates": 0}

    def process_item(self, item, spider):
        self.records.append(dict(item))
        return item

    def close_spider(self, spider):
        unique, dupes = deduplicate(self.records)
        self.stats = {"raw": len(self.records), "duplicates": len(dupes)}
        spider.logger.info("Dedup: raw=%s duplicates=%s unique=%s",
                           len(self.records), len(dupes), len(unique))
        spider.crawler.stats.set_value("cornell/raw_records", len(self.records))
        spider.crawler.stats.set_value("cornell/duplicates_removed", len(dupes))
        spider.crawler.stats.set_value("cornell/unique_firms", len(unique))
        # Records already persisted incrementally by JsonlRawStorePipeline;
        # hand the unique list to the Excel pipeline via the spider object.
        spider._unique_firms = unique


class ExcelWriterPipeline:
    """Write the final workbook: output/IP_Law_Firms.xlsx, one sheet
    'IP Law Firms', exactly two columns (Law Firm Name | Official
    Website URL)."""

    def close_spider(self, spider):
        firms = getattr(spider, "_unique_firms", [])
        out = spider.settings.get(
            "CORNELL_OUTPUT_XLSX",
            os.path.join(os.getcwd(), "output", "IP_Law_Firms.xlsx"))
        build_firm_workbook(firms, out)
        spider.logger.info("Wrote %s firm rows to %s", len(firms), out)
