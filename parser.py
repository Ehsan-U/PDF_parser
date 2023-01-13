import re
import coloredlogs
import logging
import pdfplumber
import csv
from traceback import print_exc
from pprint import pprint

class Parser():
    logger = logging.getLogger("PDF_parser")
    coloredlogs.install(level='DEBUG', logger=logger)

    def init_writer(self, filename):
        f = open(f"{filename}.csv",'w')
        writer = csv.writer(f)
        writer.writerow(['Name','Address','Phone','Website','Email','Brands','Description'])
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

    def build(self, company):
        name = company[0]
        address = self.cal_address(company)
        phone = self.cal_phone(company)
        website = self.cal_website(company)
        email = self.cal_email(company)
        brands = self.cal_brands(company)
        desc = self.cal_desc(company)

        return [name, address, phone, website, email, brands, desc]

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
        file, writer = self.init_writer('out')
        try:
            self.parse("A.pdf", writer)
        except Exception:
            print_exc()
        finally:
            file.close()



class C_parser(Parser):


    # extract companies names
    def extract_names(self, datalist):
        companies = []
        for item in datalist:
            if "Address:" in item:
                current_index = datalist.index(item)
                company_index = current_index - 1
                company_name = datalist[company_index]
                companies.append(company_name)
        return companies

    # extract company data
    def extract_company(self, datalist, companies):
        for item in companies:
            current_index = companies.index(item)
            current_company = item
            if item != companies[-1]:
                next_company = companies[current_index + 1]
                company = self.get_slice(current_company, next_company, datalist)
                yield company
            else:
                company = self.get_slice(current_company, None, datalist, last_slice=True)
    def get_slice(self, current_company, next_company, datalist, last_slice=None):
        current_company_index = datalist.index(current_company)
        if not last_slice:
            next_company_index = datalist.index(next_company)
            company = datalist[current_company_index:next_company_index]
        else:
            company = datalist[current_company_index:]
        return company


    def parse(self, pdf_file, writer):
        with pdfplumber.open(pdf_file) as pdf:
            pages = pdf.pages[1:2]
            datalist = []
            for page in pages:
                data = page.extract_text().split('\n')[1:]
                datalist += data
            companies = self.extract_names(datalist)
            for company in self.extract_company(datalist, companies):
                pprint(company)
                # resume on monday
    def main(self):
        file, writer = self.init_writer("b_test")
        try:
            self.parse("C.pdf", writer)
        except Exception:
            print_exc()
        finally:
            file.close()

c = C_parser()
c.main()