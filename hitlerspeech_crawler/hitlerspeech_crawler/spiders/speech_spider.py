# A simple scrawler for http://www.hitler.org/speeches/

import scrapy
import os

class QuotesSpider(scrapy.Spider):
    name = "hitlerspeeches"
    dl_path = "crawled_speeches"

    def start_requests(self):
        if not os.path.exists(self.dl_path):
            print("path doesn't exist. trying to make")
            os.makedirs(self.dl_path)
        else:
            print('path exists')
        urls = [
            'http://www.hitler.org/speeches'
        ]
        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):
        links = response.xpath('//a[@href]/@href').extract()
        del links[-1]
        for link in links:
            newUrl = response.urljoin(link)
            yield scrapy.Request(url=newUrl, callback=self.parseArticle)

    def parseArticle(self, response):
        title = response.xpath('//title/text()').extract()[0]
        content = ''.join(response.xpath('//blockquote//p/text()').extract())
        filename = self.dl_path + '/%s.txt' % title
        with open(filename, 'wb') as f:
            f.write(content.encode())
