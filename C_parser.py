import re
import coloredlogs
import logging
import pdfplumber
import csv
from traceback import print_exc



class C_parser():
    logger = logging.getLogger("PDF_parser")
    coloredlogs.install(level='DEBUG', logger=logger)

    def init_writer(self, filename):
        f = open(f"{filename}.csv",'w')
        writer = csv.writer(f)
        return f, writer

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
        for comp in companies:
            current_index = companies.index(comp)
            current_company = comp
            if comp != companies[-1]:
                next_company = companies[current_index + 1]
                company = self.get_slice(current_company, next_company, datalist)
                yield company
            else:
                company = self.get_slice(current_company, None, datalist, last_slice=True)
                yield company

    def get_slice(self, current_company, next_company, datalist, last_slice=None):
        current_company_index = datalist.index(current_company)
        if not last_slice:
            next_company_index = datalist.index(next_company)
            company = datalist[current_company_index:next_company_index]
        else:
            company = datalist[current_company_index:]
        return company


    def cal_address(self, company):
        for item in company:
            if "Address:" in item:
                current_index = company.index(item)
                current_value = item.split(":")[-1]
                slice = self.make_slice(current_index, current_value, company, "Tel:")
                return slice.strip()

    def cal_phone(self, company):
        phone = ''
        for item in company:
            if "Tel:" in item:
                phone = item.split(":")[-1]
        return phone.strip()

    def cal_website(self, company):
        website = ''
        for item in company:
            if "Website:" in item:
                website = item.split("Website:")[-1].strip()
        return website.strip()

    def make_slice(self, current_index, current_value, company, stopward):
        current_index += 1
        while True:
            try:
                val = company[current_index]
                if stopward in val:
                    if stopward == 'Other':
                        current_value = current_value + " " + stopward
                    break
                elif "Company" in val[:7]:
                    break
                elif ":" in val:
                    break
                else:
                    current_value = current_value + " " + val
                    current_index += 1
            except IndexError:
                break
        return current_value

    def cal_products_services(self, company):
        for item in company:
            if "Products / Services:" in item:
                current_index = company.index(item)
                current_value = item.split(":")[-1]
                slice = self.make_slice(current_index, current_value, company, "Other")
                return slice.strip()

    def cal_desc(self, company):
        for item in company:
            if "Description:" in item:
                current_index = company.index(item)
                previous_value = company[current_index-1].split("Company")[-1]
                current_value = previous_value + " " + item.split(":")[-1]
                slice = self.make_slice(current_index, current_value, company, ":")
                return slice.strip()

    def primary_contact(self, company):
        for item in company:
            if "Primary Contact:" in item:
                current_index = company.index(item)
                current_value = item.split(":")[-1]
                slice = self.make_slice(current_index, current_value, company, ":")
                return slice.strip()

    def cal_brands(self, company):
        for item in company:
            if "Brands:" in item:
                current_index = company.index(item)
                current_value = item.split(":")[-1]
                slice = self.make_slice(current_index, current_value, company, ":")
                return slice.strip()

    def build(self, company):
        name = company[0]
        address = self.cal_address(company)
        phone = self.cal_phone(company)
        website = self.cal_website(company)
        products_services = self.cal_products_services(company)
        desc = self.cal_desc(company)
        primary_contact = self.primary_contact(company)
        brands = self.cal_brands(company)
        return [name, address, phone, website, products_services, desc, primary_contact, brands]

    def parse(self, pdf_file, writer):
        with pdfplumber.open(pdf_file) as pdf:
            pages = pdf.pages[1:]
            datalist = []
            for page in pages:
                data = page.extract_text().split('\n')[1:]
                datalist += data
            companies = self.extract_names(datalist)
            for company in self.extract_company(datalist, companies):
                if company:
                    results = self.build(company)
                    self.logger.info(f" [+] {results}")
                    writer.writerow(results)

    def main(self):
        file, writer = self.init_writer("C_output")
        writer.writerow(['Name', 'Address', 'Phone', 'Website', 'Products_Services', "Description", "Primary_Contact", 'Brands'])
        try:
            self.parse("C.pdf", writer)
        except Exception:
            print_exc()
        finally:
            file.close()

c = C_parser()
c.main()