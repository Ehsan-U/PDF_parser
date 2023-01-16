import re
import coloredlogs
import logging
import pdfplumber
import csv
from traceback import print_exc

class Parser():
    logger = logging.getLogger("PDF_parser")
    coloredlogs.install(level='DEBUG', logger=logger)

    def init_writer(self, filename):
        f = open(f"{filename}.csv",'w')
        writer = csv.writer(f)
        return f, writer

    def extract_company(self, datalist):
        company = []
        for data in datalist:
            if data == ' ':
                yield company
                company.clear()
            else:
                company.append(data)

    def cal_address(self, company):
        address = ''
        for item in company:
            if 'main' in item.lower() and 'phone' in item.lower():
                return address.strip()
            else:
                address = address + " " + item
        self.logger.warning(" [+] Warning: address missing!")

    def cal_phone(self, company):
        for item in company:
            if 'main' in item.lower() and 'phone' in item.lower():
                phone = item.split(":")[-1].strip()
                return phone.strip()
        self.logger.warning(" [+] Warning: phone missing!")

    def cal_website(self, company):
        for item in company:
            if 'web' in item.lower() and 'site' in item.lower():
                website = item.split(":")[-1].strip()
                return website.strip()
        self.logger.warning(" [+] Warning: website missing!")

    def cal_email(self, company):
        for item in company:
            if "e-mail" in item.lower():
                email = item.split(":")[-1].strip()
                return email.strip()
        self.logger.warning(" [+] Warning: email missing!")

    def cal_brands(self, company):
        brands = ''
        if 'Brands' in company:
            index = company.index("Brands") + 1
            for item in company[index:]:
                brands = brands + '\n' + item
            return brands.strip()
        self.logger.warning(" [+] Warning: brands missing!")

    def cal_desc(self, company):
        desc = ''
        if 'Description' in company:
            index = company.index("Description") + 1
            for item in company[index:]:
                desc = desc + '\n' + item
            return desc.strip()
        self.logger.warning(" [+] Warning: desc missing!")

    def cal_parent(self, company):
        parent = ''
        for item in company:
            if 'Parent Company:' in item:
                index = -1
                parent = company[-1].split("Parent Company:")[-1].replace(")",'')
                return parent.strip()
        self.logger.warning(" [+] Warning: parent missing!")

    def build(self, company):
        name = company[0]
        address = self.cal_address(company)
        phone = self.cal_phone(company)
        website = self.cal_website(company)
        email = self.cal_email(company)
        brands = self.cal_brands(company)
        desc = self.cal_desc(company)
        parent = self.cal_parent(company)
        return [name, address, phone, website, email, brands, desc, parent]

    def parse(self, pdf_file, writer):
        with pdfplumber.open(pdf_file) as pdf:
            pages = pdf.pages[1:]
            datalist = []
            for page in pages:
                data = page.extract_text().replace("\xa0",' ').split('\n')
                # remove last space at each page
                if data[-1] == ' ':
                    data.pop(-1)
                datalist += data[1:]
            for company in self.extract_company(datalist):
                results = self.build(company)
                self.logger.info(f" [+] {results}")
                writer.writerow(results)

    def main(self):
        file, writer = self.init_writer('A_output')
        writer.writerow(['Name', 'Address', 'Phone', 'Website', 'Email', 'Brands', 'Description', 'Parent Company'])
        try:
            self.parse("A.pdf", writer)
        except Exception:
            print_exc()
        finally:
            file.close()

p = Parser()
p.main()