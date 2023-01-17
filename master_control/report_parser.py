import pdfplumber
import re
import csv
from traceback import print_exc
from pprint import pprint

class Parser():

    def init_writer(self):
        file = open("out.csv",'w')
        writer = csv.writer(file)
        return writer, file

    def make_slice(self, current_index, current_value, page, stopward=None):
        if stopward:
            end_index = current_index + 1
            while True:
                try:
                    val = page[end_index]
                    if stopward in val:
                        break
                    else:
                        end_index += 1
                except IndexError:
                    break
        else:
            end_index = -1
        return page[current_index:end_index]

    def find_index(self, value, page):
        for item in page:
            if value in item or item in value:
                index = page.index(item)
                return index

    def cal_name(self, slice):
        name = re.search("[A-Z]+,[A-Z]+", slice[0]).group()
        return name

    def cal_city_state_zip(self, chunk2, chunk3):
        if re.search("\d{5}", chunk2):
            chunk2 = chunk2.split(' ')
            city, state, zipcode = chunk2[0], chunk2[1], chunk2[2]
        elif re.search("\d{5}", chunk3):
            chunk3 = chunk3.split(' ')
            city, state, zipcode = chunk3[0], chunk3[1], chunk3[2]
        else:
            city, state, zipcode = '', '', ''
        return city, state, zipcode

    def cal_address(self, slice):
        if "Mailing & Home Address".lower() in "".join(slice).lower():
            chunk1 = re.search("(.*?)(?:Weekly)", slice[2]).group(1)
            chunk2 = re.search("(.*?)(?:Rate)", slice[3]).group(1)
            chunk3 = re.search("(.*?)(?:\:)", slice[4]).group(1)
            city, state, zipcode = self.cal_city_state_zip(chunk2, chunk3)
            return (chunk1, chunk2, city, state, zipcode)

    def cal_accums(self, slice):
        accums_dict = {
            "Y Gross": '',
            "Q Gross": '',
            "Y FIT": '',
            "Q FIT": '',
            "Y SS": '',
            "Q SS": '',
            "Y MED": '',
            "Q MED": '',
            "Y State 1": '',
            "Q State 1": '',
            "Y SDI": '',
            "Q SDI": '',
            "Y VFLI": '',
            "Q VFLI": '',
            "Y Local 1": '',
            "Q Local 1": '',
            "Y NYMCT TXB": '',
            "Q NYMCT TXB": '',
            "Y NYMCT TAX": '',
            "Q NYMCT TAX": '',
            "Ac 16 POP E": '',
            "Ac 21 YTD G": '',
            "Ac 22 YTD T": '',
            "Ac 36 REGUL": '',
            "Ac 37 REGUL": '',
            "Ac 38 REGUL": '',
            "Ac 39 REGUL": '',
            "Ac 40 OVERT": '',
            "Ac 41 OVERT": '',
            "Ac 42 OVERT": '',
            "Ac 43 OVERT": '',
            "Ac AU MEDIC": '',
            "Ac AV MEDIC": '',
            "Ac AW DENTA": '',
            "Ac AX DENTA": '',
            "Ac AY VISIO": '',
            "Ac A1 VISIO": '',
            "Ac 92 PAYCH": '',
            "Ac 93 PAYCH": '',
            "Ac 0J HolDi": '',
            "Ac 0K HolDi": '',
            "Ac 0L HolDi": '',
            "Ac 0M HolDi": '',
        }
        for acc in accums_dict.keys():
            for s in slice:
                if acc in s:
                    s_regx = re.search(r"(?:.*?)(?:HEALTH|VISION|Deposits|CodeCK1|Deposit|Exemptions|S-Single|SUI/DI|\sFLI|CIT|Local|Available)(.*)", s)
                    if s_regx:
                        s_val = s_regx.group(1).strip()
                        value = re.search(rf"(?:Local\s1|)([0-9\s]+?)(?:{acc})", s_val).group(1).strip()
                    else:
                        value = re.search(rf"(?:Local\s1|)([0-9\s]+?)(?:{acc})", s).group(1).strip()
                    if value.count(' ') == 1:
                        accums_dict[acc] = value.replace(' ','.')
                    elif value.count(' ') == 2:
                        accums_dict[acc] = value.replace(' ', '',1).replace(' ','.')
        pprint(accums_dict)

    def extract_person_data(self, names, page):
        page = page.extract_text().split('\n')
        for name in names:
            current_index = self.find_index(name, page)
            current_value = name
            if names[-1] == name:
                slice = self.make_slice(current_index, current_value, page)
            else:
                next_value = names[names.index(name) + 1]
                slice = self.make_slice(current_index, current_value, page, next_value)
            # name = self.cal_name(slice)
            # address = self.cal_address(slice)
            print(name)
            print(self.cal_accums(slice))
            yield None

    def parse(self, pdf):
        with pdfplumber.open(pdf) as p:
            page =p.pages[1]
            names = [item.get('text') for item in page.search("[A-Z]+,[A-Z]+", regex=True)]
            for person in self.extract_person_data(names, page):
                pass

    def main(self):
        writer, file = self.init_writer()
        try:
            self.parse('report.pdf')
        except Exception:
            print_exc()
        else:
            file.close()



p = Parser()
p.main()