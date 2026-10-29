import scrapy


class FirmItem(scrapy.Item):
    firm_name = scrapy.Field()
    website = scrapy.Field()
    state = scrapy.Field()
    city = scrapy.Field()
    # provenance / audit fields (not exported to Excel)
    lawyer_name = scrapy.Field()
    profile_url = scrapy.Field()
    listing_url = scrapy.Field()
