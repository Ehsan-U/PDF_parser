import pdfplumber
import re
import csv
from traceback import print_exc
from pprint import pprint

class Parser():
    def __init__(self):
        self.header = True
        self.last_person = ''
        self.persons = []

    def init_writer(self):
        file = open("out.csv",'w', newline='')
        self.writer = csv.writer(file)
        return file

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
        try:
            name = re.search("[A-Z]+,[A-Z]+", slice[0]).group()
        except AttributeError:
            name = re.search("(.*?)(?:\sGross\:)", slice[0]).group(1)
        return name

    def cal_city_state_zip(self, addr2, addr3):
        if re.search("\d{5}", addr2):
            addr2 = addr2.split(' ')
            if len(addr2) == 5:
                city, state, zipcode = addr2[0] + ' ' + addr2[1] + ' ' + addr2[2], addr2[3], addr2[-1]
            elif len(addr2) == 4:
                city, state, zipcode = addr2[0]+ ' ' +addr2[1], addr2[2], addr2[-1]
            else:
                city, state, zipcode = addr2[0] ,addr2[1], addr2[-1]
        elif re.search("\d{5}", addr3):
            addr3 = addr3.split(' ')
            if len(addr3) == 5:
                city, state, zipcode = addr3[0] + ' ' + addr3[1] + ' ' + addr3[2], addr3[3], addr3[-1]
            elif len(addr3) == 4:
                city, state, zipcode = addr3[0]+ ' ' +addr3[1], addr3[2], addr3[-1]
            else:
                city, state, zipcode = addr3[0] ,addr3[1], addr3[-1]
        else:
            city, state, zipcode = '', '', ''
        return city, state, zipcode

    def cal_address(self, slice):
        addr1, addr2, city, state, zipcode = '', '', '', '', ''
        try:
            if "Mailing & Home Address".lower() in "".join(slice).lower():
                addr1 = re.search("(.*?)(?:Weekly)", slice[2]).group(1).strip()
                addr2 = re.search("(.*?)(?:Rate)", slice[3]).group(1).replace("LWW",'').strip()
                addr3 = re.search("(.*?)(?:\:|Rate)", slice[4]).group(1).replace("LWW",'').strip()
                city, state, zipcode = self.cal_city_state_zip(addr2, addr3)
        except Exception:
            pass
        return (addr1, addr2, city, state, zipcode)

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
                    try:
                        s_regx = re.search(r"(?:.*?)(?:HEALTH|VISION|Deposits|CodeCK1|Deposit|Exemptions|S-Single|SUI/DI|\sFLI|CIT|Local|Available)(.*)", s)
                        if s_regx:
                            s_val = s_regx.group(1).strip()
                            value = re.search(rf"(?:\w+\s1|)(\s[0-9\s]+?)(?:{acc})", s_val).group(1).strip()
                        else:
                            value = re.search(rf"(?:\w+\s1|)(\s[0-9\s]+?)(?:{acc})", s).group(1).strip()
                        if value.count(' ') == 1:
                            accums_dict[acc] = value.replace(' ','.')
                        elif value.count(' ') == 2:
                            accums_dict[acc] = value.replace(' ', '',1).replace(' ','.')
                    except AttributeError:
                        pass
        return accums_dict


    def organize(self, name, address, accumulations):
        data = {
            "Name":name,
            "Address1":address[0],
            "Address2": address[1],
            "City": address[2],
            "State": address[3],
            "Zip": address[4],
        }
        for k, v in accumulations.items():
            data[k] = v
        if self.header:
            self.header = False
            self.writer.writerow(data.keys())
        return data

    def extract_person_data(self, names, page):
        result = ''
        page = page.extract_text().split('\n')
        for name in names:
            current_index = self.find_index(name, page)
            if current_index != None:
                current_value = name
                if names[-1] == name:
                    slice = self.make_slice(current_index, current_value, page)
                else:
                    next_value = names[names.index(name) + 1]
                    slice = self.make_slice(current_index, current_value, page, next_value)
                name = self.cal_name(slice)
                address = self.cal_address(slice)
                accumulations = self.cal_accums(slice)
                result = self.organize(name, address, accumulations)
            yield result


    def merge_dicts(self, old_dict, new_dict):
        for key,val in new_dict.items():

            if old_dict[key]:
                if val and key != 'Name' and old_dict[key] != new_dict[key]:
<<<<<<< HEAD
                    old_dict[key] = old_dict[key] + " " + val
=======
                    print(f"{key=}, {old_dict['Name']=},{new_dict['Name']=}, {old_dict[key]=}, {new_dict[key]=}")
                pass
>>>>>>> 043a2bf917994e8f44ffdec3877af8dc1e61d796
            else:
                old_dict[key] = val
        return old_dict

    def is_exist(self, person):
        for p in self.persons:
            if person['Name'] in p:
                old_dict = p.get(person['Name'])
                new_dict = person
                updated = self.merge_dicts(old_dict, new_dict)
                p[person['Name']] = updated
                return True


    def extract_person_names(self, page):
        names = []
        regx = re.findall("(.*?)(?:\sGross\:)|([A-Z]+,[A-Z]+)", page.extract_text())
        for name in regx:
            if any(name):
                name = "".join([n for n in name if n])
                names.append(name)
        return names

    def parse(self, pdf):
        with pdfplumber.open(pdf) as p:
            for n, page in enumerate(p.pages[7:-1], start=1):
                print(f"\r [+] Parsed pages: {n}",end='')
                names = self.extract_person_names(page)
                for person in self.extract_person_data(names, page):
                    if person:
                        if self.is_exist(person):
                            continue
                        self.persons.append({person['Name']: person})

            for person in self.persons:
                value = list(person.values())[0]
                self.writer.writerow(value.values())

    def main(self):
        file = self.init_writer()
        try:
            self.parse('file.pdf')
        except Exception:
            print_exc()
        else:
            file.close()


p = Parser()
p.main()
