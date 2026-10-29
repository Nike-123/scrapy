BOT_NAME = "cornell_ip"
SPIDER_MODULES = ["cornell_ip.spiders"]
NEWSPIDER_MODULE = "cornell_ip.spiders"
ROBOTSTXT_OBEY = True
USER_AGENT = "cornell-ip-firms-research (+https://lawyers.law.cornell.edu)"
CONCURRENT_REQUESTS = 8
DOWNLOAD_DELAY = 0.5
AUTOTHROTTLE_ENABLED = True
RETRY_TIMES = 3
LOG_LEVEL = "INFO"
# Final two-column workbook and resumable data store.  Run scrapy from
# this project directory (or override with -s), and pass
#   -s JOBDIR=data/crawljob
# to pause/resume a crawl with Scrapy's persistent scheduler queue.
CORNELL_DATA_DIR = "data"
CORNELL_OUTPUT_XLSX = "output/IP_Law_Firms.xlsx"
