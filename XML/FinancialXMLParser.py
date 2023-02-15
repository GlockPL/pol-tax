import xml.etree.ElementTree as ET
import pandas as pd
from urllib.parse import unquote
from urllib.parse import urlparse
from pathlib import Path
import xlsxwriter
import argparse
import re
import base64

from decimal import *


# My files Imports
from XML.xsd_downloader import download_schema, XSD_DIR
from Settings import Settings


class FinancialXMLParser:
    def __init__(self, xml_path: str, app_settings: Settings):
        self.app_settings = app_settings
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
        self.xsd_dict.update(app_settings.get_keywords())
        # get introduction
        self.introduction_dict = self.find_introduction()
        self.introduction_pd, self.introduction_border_list = self.generate_introduction_pd()
        # get bilans
        self.financial_balance_dict = self.find_bilans_and_parse()
        # get aktywa dict
        self.assets_dict = self.get_assets()
        # get pasywa dict
        self.passives_dict = self.get_equity_plus_liabilities()
        # Embedde file extraction
        self.embedded_files_dict = self.get_embedded_files()
        # Aktywa pd
        self.assets_pd, self.assets_border_list = self.generate_assets_pd()
        # Pasywa pd
        self.passives_pd, self.passives_border_list = self.generate_passives_pd()

    def list_all_root_tags(self):
        for elem in self.root:
            print(elem.tag)

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
                        val = re.sub('\[.*\]', '', val)
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

    def xml_tag_to_dict(self, xml_tag, xml_dict):
        tag = xml_tag.tag
        lastChar = tag.find('}')
        if lastChar > 0:
            tag = tag[lastChar + 1:]

        text = xml_tag.text.strip() if xml_tag.text is not None else ""
        if text == "":
            xml_dict[tag] = {}
        else:
            xml_dict[tag] = text

        for elem in xml_tag:
            xml_dict[tag] = self.xml_tag_to_dict(elem, xml_dict[tag])

        return xml_dict

    def find_bilans_and_parse(self):
        bilans_tag_elem = self.recurse_search(self.root, 'Bilans')
        bilans_dict_xml = {}
        bilans_dict_xml = self.xml_tag_to_dict(bilans_tag_elem, bilans_dict_xml)
        return bilans_dict_xml

    def find_introduction(self):
        # WprowadzenieDoSprawozdaniaFinansowego
        introduction_tag_elem = self.recurse_search(self.root, 'WprowadzenieDoSprawozdaniaFinansowego')
        introduction_dict_xml = {}
        introduction_dict_xml = self.xml_tag_to_dict(introduction_tag_elem, introduction_dict_xml)

        ret_dict = {'WprowadzenieDoSprawozdaniaFinansowego': {}}
        for k, v in introduction_dict_xml['WprowadzenieDoSprawozdaniaFinansowego'].items():
            match = re.findall('P_([0-9]{1,3})', k)
            if len(match) > 0 and int(match[0]) < 4:
                ret_dict['WprowadzenieDoSprawozdaniaFinansowego'][k] = v

        return ret_dict

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

    def get_embedded_files(self):
        embedded_tag_elem = self.recurse_search(self.root, 'DodatkoweInformacjeIObjasnienia')
        embedded_dict_xml = {}
        embedded_dict_xml = self.xml_tag_to_dict(embedded_tag_elem, embedded_dict_xml)
        return embedded_dict_xml

    def get_assets(self):
        level = 0
        # aktywa = defaultdict(list)
        aktywa = {}
        for k, v in self.financial_balance_dict['Bilans']['Aktywa'].items():
            aktywa = self.travel_dict_reccur(k, v, aktywa, 'Aktywa', 'Aktywa', None, level)

        recombined_aktywa = self.join_two_subtags("KwotaA", "KwotaB", "Aktywa", aktywa)
        return recombined_aktywa

    def get_introduction(self):
        level = 0
        introduction = {}
        for k, v in self.introduction_dict.items():
            introduction = self.travel_dict_reccur(k, v, introduction, 'WprowadzenieDoSprawozdaniaFinansowego', 'WprowadzenieDoSprawozdaniaFinansowego', None, level)

        return introduction

    def get_equity_plus_liabilities(self):
        level = 0
        passives = {}
        for k, v in self.financial_balance_dict['Bilans']['Pasywa'].items():
            passives = self.travel_dict_reccur(k, v, passives, 'Pasywa', 'Pasywa', None, level)

        recombined_passives = self.join_two_subtags("KwotaA", "KwotaB", "Pasywa", passives)
        return recombined_passives

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

                    temp_dict[cost_type] = Decimal(cost)#float(cost)
                new_dict[main_tag][k] = new_list
            elif float(v) is not None:
                new_dict[main_tag]["Koszta"][k] = Decimal(v)#float(v)
        return new_dict

    def col_size(self, string):
        return len(str(string)) - 1

    def check_cols_max_size(self, max_list, keys):
        for idx, val in enumerate(keys):
            key_size = self.col_size(val)
            if max_list[idx] < key_size:
                max_list[idx] = key_size

        return max_list

    def reccur_flatten(self, reccur_obj, upper_tag, return_dict: dict):
        if type(reccur_obj) == dict:
            for k, v in reccur_obj.items():
                if type(v) == dict:
                    return_dict[k] = "Header"
                    self.reccur_flatten(v, k, return_dict)
                else:
                    return_dict[k] = v
        else:
            return_dict[upper_tag] = return_dict

        return return_dict

    def flatten_dict(self, dict_to_flatten: dict):
        out_dict = {}

        for k, v in dict_to_flatten.items():
            out_dict = self.reccur_flatten(v, k, out_dict)

        return out_dict

    def generate_introduction_pd(self):
        df = pd.DataFrame(columns=["Col1", "Col2", "Col3"])
        border_this_list = []
        flatten_introduction = self.flatten_dict(self.introduction_dict)
        row = 0
        for k, v in flatten_introduction.items():
            try:
                row_name = self.xsd_dict[k]
            except KeyError:
                row_name = k

            if row_name == "":
                continue

            if v == "Header":
                val = ""
                pd_row = [row_name, "", ""]
                border_this_list.append(1)
            else:
                val = v
                pd_row = ["", row_name, val]
                border_this_list.append(0)

            df.loc[row] = pd_row

            row += 1

        return df, border_this_list

    def color_bg_row(self, s, border_list):
        if border_list[s.index] == 1:
            return ['background-color: #d7e4bc']

        return ['']

    def generate_introduction_excel_from_pd(self, out_file_name):
        workbook = xlsxwriter.Workbook(out_file_name)
        worksheet = workbook.add_worksheet()
        workbook, header_cell_format, normal_cell_format = self.excel_style_format(workbook)

        col_max_size = [0, 0]
        i = 0
        for index, row in self.introduction_pd.iterrows():
            if self.introduction_border_list[i] == 1:
                worksheet.write(i, 0, row['Col1'], header_cell_format)
                col_max_size = self.check_cols_max_size(col_max_size, [row['Col1'], ""])
            else:
                worksheet.write(i, 0, row['Col2'], normal_cell_format)
                worksheet.write(i, 1, row['Col3'], normal_cell_format)
                col_max_size = self.check_cols_max_size(col_max_size, [row['Col2'], row['Col3']])
            i += 1

        for idx, col_siz in enumerate(col_max_size):
            worksheet.set_column(idx, idx, col_siz)

        workbook.close()

    def excel_style_format(self, workbook):
        header_cell_format = workbook.add_format()
        header_cell_format.set_pattern(1)  # This is optional when using a solid fill.
        header_cell_format.set_bg_color('#d7e4bc')
        header_cell_format.set_border(1)

        normal_cell_format = workbook.add_format()
        normal_cell_format.set_pattern(1)  # This is optional when using a solid fill.
        normal_cell_format.set_bg_color('#FEFEFE')

        return workbook, header_cell_format, normal_cell_format

    def generate_introduction_excel(self, out_file_name):
        workbook = xlsxwriter.Workbook(out_file_name)
        worksheet = workbook.add_worksheet()

        header_cell_format = workbook.add_format()
        header_cell_format.set_pattern(1)  # This is optional when using a solid fill.
        header_cell_format.set_bg_color('#d7e4bc')
        header_cell_format.set_border(1)

        normal_cell_format = workbook.add_format()
        normal_cell_format.set_pattern(1)  # This is optional when using a solid fill.
        normal_cell_format.set_bg_color('#FEFEFE')

        flatten_introduction = self.flatten_dict(self.introduction_dict)
        row = 0
        col_max_size = [0, 0]
        for k, v in flatten_introduction.items():
            try:
                row_name = self.xsd_dict[k]
            except KeyError:
                row_name = k

            if row_name == "":
                continue

            col_max_size = self.check_cols_max_size(col_max_size, [row_name, v])

            if v == "Header":
                cell_format = header_cell_format
                val = ""
                pd_row = [row_name, "", ""]
            else:
                cell_format = normal_cell_format
                val = v
                pd_row = ["", row_name, val]

            worksheet.write(row, 0, row_name, cell_format)
            worksheet.write(row, 1, val, cell_format)
            self.introduction_pd.loc[row] = pd_row

            row += 1

        for idx, col_siz in enumerate(col_max_size):
            worksheet.set_column(idx, idx, col_siz)
        workbook.close()

    def generate_passives_excel(self, out_file_name):
        self.generate_excel(out_file_name, self.passives_dict, base_tag="Pasywa")

    def generate_assets_excel(self, out_file_name):
        self.generate_excel(out_file_name, self.assets_dict, base_tag="Aktywa")

    def generate_assets_pd(self):
        return self.generate_costs_pd(self.assets_dict, base_tag="Aktywa")

    def generate_passives_pd(self):
        return self.generate_costs_pd(self.passives_dict, base_tag="Pasywa")

    def generate_passive_pd(self):
        return self.generate_costs_pd(self.passives_dict, base_tag="Pasywa")

    def generate_costs_pd(self, main_dict, base_tag="Aktywa"):
        border_this_list = []
        temp_dict = main_dict[base_tag].copy()
        df = pd.DataFrame(columns=["", self.xsd_dict[self.cost_tag_1], self.xsd_dict[self.cost_tag_2]])

        row = 0
        df.loc[row+1] = [self.xsd_dict[base_tag], temp_dict['Koszta'][self.cost_tag_1],
                       temp_dict['Koszta'][self.cost_tag_2]]
        border_this_list.append(0)

        row += 1
        del temp_dict['Koszta']

        border_count = 0
        for k, v in temp_dict.items():
            if type(v) == list:
                prefix = k.replace(f"{base_tag}_", "")
                df.loc[row+1] = [f"{prefix}. {self.xsd_dict[k]}", v[0]['Koszta'][self.cost_tag_1],
                               v[0]['Koszta'][self.cost_tag_2]]
                border_this_list.append(1)
                row += 1
                for d in v[1:]:
                    for list_k, list_v in d.items():
                        if border_count % 3 == 0:
                            border_this_list.append(1)
                        else:
                            border_this_list.append(0)

                        prefix = list_k.replace(f"{base_tag}_", "").replace("_", ".")
                        df.loc[row+1] = [f"{prefix} {self.xsd_dict[list_k]}", list_v[self.cost_tag_1],
                                       list_v[self.cost_tag_2]]
                        row += 1
                        border_count += 1
        return df, border_this_list

    def generate_excel(self, out_file_name, main_dict, base_tag="Aktywa"):
        workbook = xlsxwriter.Workbook(out_file_name)
        worksheet = workbook.add_worksheet()
        number_format = workbook.add_format({'num_format': '#,##0.00'})
        cell_format = workbook.add_format()
        cell_format.set_pattern(1)  # This is optional when using a solid fill.
        cell_format.set_bg_color('#d7e4bc')
        cell_format.set_border(1)

        cell_format_thicker = workbook.add_format()
        cell_format_thicker.set_pattern(1)  # This is optional when using a solid fill.
        cell_format_thicker.set_bg_color('#d7e4bc')
        cell_format_thicker.set_border(1)
        cell_format_thicker.set_bottom(2)

        temp_dict = main_dict[base_tag].copy()

        row = 0
        col = 0
        worksheet.write(row, col + 1, self.xsd_dict[self.cost_tag_1])
        worksheet.write(row, col + 2, self.xsd_dict[self.cost_tag_2])
        row += 1
        worksheet.write(row, col, self.xsd_dict[base_tag], cell_format)
        worksheet.write_number(row, col + 1, temp_dict['Koszta'][self.cost_tag_1], number_format)
        worksheet.write_number(row, col + 2, temp_dict['Koszta'][self.cost_tag_2], number_format)
        row += 1
        del temp_dict['Koszta']
        col_max_size = [0, 0, 0]
        col_max_size = self.check_cols_max_size(col_max_size,
                                                [self.xsd_dict[base_tag], self.xsd_dict[self.cost_tag_1],
                                                 self.xsd_dict[self.cost_tag_2]])
        border_count = 0
        for k, v in temp_dict.items():
            if type(v) == list:
                prefix = k.replace(f"{base_tag}_", "")
                worksheet.write(row, col, f"{prefix}. {self.xsd_dict[k]}", cell_format_thicker)
                worksheet.write_number(row, col + 1, v[0]['Koszta'][self.cost_tag_1], number_format)
                worksheet.write_number(row, col + 2, v[0]['Koszta'][self.cost_tag_2], number_format)
                row += 1
                border_count += 1
                for d in v[1:]:
                    for list_k, list_v in d.items():
                        if border_count % 3 == 0:
                            formt = cell_format_thicker
                        else:
                            formt = cell_format

                        col_max_size = self.check_cols_max_size(col_max_size,
                                                                [self.xsd_dict[list_k], list_v[self.cost_tag_1],
                                                                 list_v[self.cost_tag_2]])
                        prefix = list_k.replace(f"{base_tag}_", "").replace("_", ".")
                        worksheet.write(row, col, f"{prefix} {self.xsd_dict[list_k]}", formt)
                        worksheet.write_number(row, col + 1, list_v[self.cost_tag_1], number_format)
                        worksheet.write_number(row, col + 2, list_v[self.cost_tag_2], number_format)
                        row += 1
                        border_count += 1

        for idx, col_siz in enumerate(col_max_size):
            worksheet.set_column(idx, idx, col_siz)

        workbook.close()

    def extract_embedded_files(self):
        files_dict = self.embedded_files_dict['DodatkoweInformacjeIObjasnieniaJednostkaInna']['DodatkoweInformacjeIObjasnienia']
        files_list = []
        for k, v in files_dict.items():
            if k == "Plik":
                files_list.append({'file_name': v['Nazwa'], 'content': v['Zawartosc']})

        for elem in files_list:
            file_path = Path(elem['file_name'])
            concatenated_path = Path.joinpath(self.app_settings.get_extracted_dir(), file_path)
            base64_bytes = elem['content'].encode('utf8')
            message_bytes = base64.b64decode(base64_bytes)

            with concatenated_path.open(mode='wb') as fh:
                fh.write(message_bytes)

        return concatenated_path


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
        # f_xml.generate_excel('Aktywa.xlsx')
        f_xml.generate_assets_excel('Aktywa.xlsx')
        f_xml.generate_passives_excel('Pasywa.xlsx')
        # passives = f_xml.get_equity_plus_liabilities()
        # print(passives)
    # print(f_xml.financial_balance_dict)
