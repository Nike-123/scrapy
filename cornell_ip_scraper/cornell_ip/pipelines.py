"""Pipelines: deduplication (first occurrence wins) + Excel writer."""
import os

from cornell_ip.dedup import deduplicate
from cornell_ip.excel import build_workbook


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
        # Hand the unique list to the Excel pipeline via the spider object.
        spider._unique_firms = unique


class ExcelWriterPipeline:
    def close_spider(self, spider):
        firms = getattr(spider, "_unique_firms", [])
        out = spider.settings.get(
            "CORNELL_OUTPUT_XLSX",
            os.path.join(os.getcwd(), "US_IP_Law_Firms_Cornell.xlsx"))
        build_workbook(firms, out)
        spider.logger.info("Wrote %s firm rows to %s", len(firms), out)
