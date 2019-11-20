from urllib.parse import urlsplit
from lxml.etree import XMLSyntaxError
from itertools import chain
import lxml.html as lxml_parser
import logging
import requests
import xlsxwriter

_logger = logging.getLogger().setLevel(logging.DEBUG)


class Parser:

    def __init__(self):
        self._url = "https://setam.net.ua/neruhomist/zemlya/filters/state=102"
        self.x_delimiter = [("ХХХХХХХ", "ХХХХХХХ")]
        self.number_delimiter = [("11111111", "11111111")]
        self.city_delimiter = [("Київ", "Київ")]
        self.pdv_delimiter = [("без ПДВ", "без ПДВ")]
        self.currency_delimiter = [("Гривня", "Гривня")]
        self.text = 'Текст'
        self.link_text = 'Ссылка на страницу'

    def get_domain(self, link):
        return "{0.scheme}://{0.netloc}/".format(urlsplit(link))

    def get_url_content(self, url):
        try:
            response = requests.get(url)
            return response.text
        except requests.exceptions.RequestException as ex:
            _logger.debug(f"{ex} while getting response")

    def get_tree(self, html):
        try:
            return lxml_parser.fromstring(html)
        except (XMLSyntaxError,
                TypeError,
                ValueError) as ex:
            _logger.debug(f"Exception while getting html tree {ex}")

    def get_html(self, url):
        content = self.get_url_content(url)
        return self.get_tree(content)

    @staticmethod
    def strip_text(data):
        stripped_data = [el.strip() for el in data if el.strip()]
        return stripped_data[0], stripped_data[1]

    def parse_inner_pages(self, link):
        extracted_data = []
        page_content = self.get_html(link)
        base_path = page_content.xpath("//div[@class='panel-body']/div[position()>1]")
        for n in range(len(base_path) - 1):
            data_info = [self.strip_text(tag.xpath(".//text()"))
                         for tag in base_path[n].xpath(".//div[@class='date-end-row'][position()<3 or position()>3]")]
            price_info = [self.strip_text(base_path[n].xpath(".//div[@class='start-price-row']//text()"))]
            publicity_date = [self.strip_text(base_path[n].xpath(".//div[@class='payment-row'][last()]//text()"))]
            text = [(self.text, " ".join(
                [sent.strip() for sent in base_path[n + 1].xpath(".//div[@id='Feature-lot']//p/text()")]))]
            link = [(self.link_text, link)]
            extracted_data = list(chain(link, self.x_delimiter, data_info, self.x_delimiter,
                                        self.number_delimiter, self.city_delimiter, self.city_delimiter,
                                        text, self.x_delimiter, self.x_delimiter, self.x_delimiter,
                                        price_info, self.pdv_delimiter, self.currency_delimiter, publicity_date))
        return extracted_data

    def parse(self):
        results = []
        tree = self.get_html(self._url)
        articles = tree.xpath("//div[@class = 'tab-content']/descendant::div[1]/div[position()=1]/*")
        if not articles:
            return results
        for article in articles:
            url = article.xpath(".//a/@href")[0]
            url = url if url.startswith("http") else self.get_domain(self._url) + url
            results.append(self.parse_inner_pages(url))
        return results

    def write_to_excel(self):
        workbook = xlsxwriter.Workbook('results.xlsx')
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': 1})
        data = self.parse()
        row = 0
        self.write_columns_names(bold, data, row, worksheet)
        self.write_data_to_rows(data, row, worksheet)
        workbook.close()

    def write_data_to_rows(self, data, row, worksheet):
        for el in data:
            row = row + 1
            for i, line in enumerate(el):
                worksheet.write(row, i, line[1])

    def write_columns_names(self, marker, data, row, worksheet):
        for el in data:
            for i, line in enumerate(el):
                worksheet.write(row, i, line[0], marker)


if __name__ == '__main__':
    Parser().write_to_excel()
