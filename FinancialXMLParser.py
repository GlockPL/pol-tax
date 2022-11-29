import xml.etree.ElementTree as ET
from urllib.parse import unquote
from urllib.parse import urlparse
from pathlib import Path
from xsd_downloader import download_schema, XSD_DIR
import xlsxwriter
import argparse


class FinancialXMLParser:
    def __init__(self, xml_path):
        self.cost_tag_1 = "KwotaA"
        self.cost_tag_2 = "KwotaB"
        self.xml_path = xml_path
        self.tree = ET.parse(xml_path)
        self.root = self.tree.getroot()
        # Creates cache dir
        self.cache_dir = XSD_DIR
        Path(self.cache_dir).mkdir(parents=True, exist_ok=True)
        # Gets schema url from xml file
        self.schema_url = self.get_schema_url()
        self.download_xsd()
        self.xsd_dict = self.map_xsd()
        # get bilans
        self.financial_balance_dict = self.find_bilans_and_parse()
        # get aktywa dict
        self.assets_dict = self.get_assets()

    def remove_curly_braces(self, name):
        last_char = name.find('}')
        if last_char > 0:
            name = name[last_char + 1:]
        return name

    def get_schema_url(self):
        schema_url = None
        for elem_k, elem_v in self.root.attrib.items():
            if self.remove_curly_braces(elem_k) == "schemaLocation":
                last_ocr = elem_v.rfind("http")
                schema_url = unquote(elem_v[last_ocr:])
                parse = urlparse(schema_url)
                if parse.netloc == 'www.mf.gov.pl':
                    schema_url = parse._replace(netloc='www.gov.pl')._replace(scheme='https').geturl()
                    # parse.netloc =
                    # parse.scheme = 'https'

        return schema_url

    def download_xsd(self):
        download_schema(self.schema_url)

    def recur_annot(self, root, dict_obj):
        for child in root:
            if "element" in child.tag:
                name = child.attrib['name']
                for elem in child:
                    if "annotation" in elem.tag:
                        val = elem[0].text
                        # print(f"{name}:{val}")
                        dict_obj[name] = val
            self.recur_annot(child, dict_obj)

    def map_xsd(self):
        p = Path(self.cache_dir)
        xsd_files = p.glob('*.xsd')
        xsd_map = {}
        for file in xsd_files:
            tree = ET.parse(file)
            root = tree.getroot()
            self.recur_annot(root, xsd_map)

        return xsd_map

    def recurse_search(self, root, tag):
        for elem in root:
            if tag in elem.tag:
                return elem
            else:
                self.recurse_search(elem, tag)

    def bilans_parsing(self, bilans_tag, bilans_dict):
        tag = bilans_tag.tag
        lastChar = tag.find('}')
        if lastChar > 0:
            tag = tag[lastChar + 1:]

        text = bilans_tag.text.strip() if bilans_tag.text is not None else ""
        if text == "":
            bilans_dict[tag] = {}
        else:
            bilans_dict[tag] = text

        for elem in bilans_tag:
            bilans_dict[tag] = self.bilans_parsing(elem, bilans_dict[tag])

        return bilans_dict

    def find_bilans_and_parse(self):
        bilans_tag_elem = self.recurse_search(self.root, 'Bilans')
        bilans_dict_xml = {}
        bilans_dict_xml = self.bilans_parsing(bilans_tag_elem, bilans_dict_xml)
        return bilans_dict_xml

    def travel_dict_reccur(self, key, value, dict_to_fill, upper_key, level_0_key=None, level_1_key=None, level=0):
        if type(value) is dict:
            for k, v in value.items():
                if level == 0:
                    level_1_key = key
                dict_to_fill = self.travel_dict_reccur(k, v, dict_to_fill, key, level_0_key, level_1_key, level + 1)
        else:
            if level == 0:
                if level_0_key not in dict_to_fill:
                    dict_to_fill[level_0_key] = {}
                dict_to_fill[level_0_key][key] = value
            elif level == 1:
                if level_1_key not in dict_to_fill[level_0_key]:
                    dict_to_fill[level_0_key][level_1_key] = []
                dict_to_fill[level_0_key][level_1_key].append((upper_key, key, value))
            else:
                dict_to_fill[level_0_key][level_1_key].append((upper_key, key, value))

        return dict_to_fill

    def get_assets(self):
        level = 0
        # aktywa = defaultdict(list)
        aktywa = {}
        for k, v in self.financial_balance_dict['Bilans']['Aktywa'].items():
            aktywa = self.travel_dict_reccur(k, v, aktywa, 'Aktywa', 'Aktywa', None, level)

        recombined_aktywa = self.join_two_subtags("KwotaA", "KwotaB", "Aktywa", aktywa)
        # print(recombined_aktywa)
        return recombined_aktywa

    # TODO: Change float to decimal
    def join_two_subtags(self, tag1, tag2, main_tag, data_dict):
        new_dict = {main_tag: {"Koszta": {}}}
        for k, v in data_dict[main_tag].items():
            if type(v) == list:
                last_tag = None
                temp_dict = {}
                new_list = []
                for tag, cost_type, cost in v:
                    if tag != last_tag:
                        temp_dict = {}
                        last_tag = tag
                    else:
                        last_tag = None
                        tag_to_write = tag if k != tag else "Koszta"
                        new_list.append({tag_to_write: temp_dict})

                    temp_dict[cost_type] = float(cost)
                new_dict[main_tag][k] = new_list
            elif float(v) is not None:
                new_dict[main_tag]["Koszta"][k] = float(v)
        return new_dict

    def col_size(self, string):
        return len(str(string)) - 1

    def check_cols_max_size(self, max_list, keys):
        for idx, val in enumerate(keys):
            key_size = self.col_size(val)
            if max_list[idx] < key_size:
                max_list[idx] = key_size

        return max_list

    def generate_excel(self, out_file_name):
        workbook = xlsxwriter.Workbook(out_file_name)
        worksheet = workbook.add_worksheet()
        cell_format = workbook.add_format()
        cell_format.set_pattern(1)  # This is optional when using a solid fill.
        cell_format.set_bg_color('#c2bebe')

        temp_dict = self.assets_dict['Aktywa'].copy()

        row = 0
        col = 0
        worksheet.write(row, col + 1, self.xsd_dict[self.cost_tag_1])
        worksheet.write(row, col + 2, self.xsd_dict[self.cost_tag_2])
        row += 1
        worksheet.write(row, col, self.xsd_dict["Aktywa"], cell_format)
        worksheet.write_number(row, col + 1, temp_dict['Koszta'][self.cost_tag_1], cell_format)
        worksheet.write_number(row, col + 2, temp_dict['Koszta'][self.cost_tag_2], cell_format)
        row += 1
        del temp_dict['Koszta']
        col_max_size = [0, 0, 0]
        col_max_size = self.check_cols_max_size(col_max_size,
                                                [self.xsd_dict["Aktywa"], self.xsd_dict[self.cost_tag_1],
                                                 self.xsd_dict[self.cost_tag_2]])

        for k, v in temp_dict.items():
            if type(v) == list:
                worksheet.write(row, col, self.xsd_dict[k], cell_format)
                worksheet.write_number(row, col + 1, v[0]['Koszta'][self.cost_tag_1], cell_format)
                worksheet.write_number(row, col + 2, v[0]['Koszta'][self.cost_tag_2], cell_format)
                row += 1
                for d in v[1:]:
                    for list_k, list_v in d.items():
                        col_max_size = self.check_cols_max_size(col_max_size,
                                                                [self.xsd_dict[list_k], list_v[self.cost_tag_1],
                                                                 list_v[self.cost_tag_2]])
                        worksheet.write(row, col, self.xsd_dict[list_k])
                        worksheet.write_number(row, col + 1, list_v[self.cost_tag_1])
                        worksheet.write_number(row, col + 2, list_v[self.cost_tag_2])
                        row += 1

        for idx, col_siz in enumerate(col_max_size):
            worksheet.set_column(idx, idx, col_siz)

        workbook.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Program do kowersji xmlów z KRSu na pliki excel")
    parser.add_argument(
        "-f", "--file", type=str, help='Ścieżka do pliku xml'
    )
    args = parser.parse_args()
    path = "./01_Coca-Cola_FS_2021_v6_30.06.2022.xml"
    path = "./P300_2020_20220214.xml"
    path = "./PEPSI-COLA GENERAL BOTTLERS POLAND Sp z oo SprFinJednostkaInnaWTysiacach 2022.07.28.xml"
    path = args.file

    if path is not None:
        f_xml = FinancialXMLParser(path)
        f_xml.generate_excel('Aktywa.xlsx')
    # print(f_xml.financial_balance_dict)
