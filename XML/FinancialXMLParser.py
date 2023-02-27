import xml.etree.ElementTree as ET
import pandas as pd
from urllib.parse import unquote
from urllib.parse import urlparse
from pathlib import Path
import xlsxwriter
import argparse
import re
import base64
import logging

from decimal import *

# My files Import
from XML.xsd_downloader import download_schema, XSD_DIR
from Settings import Settings, loging_level

logging.basicConfig(level=loging_level)


def excel_style_format(workbook):
    font_type = "Arial"
    font_size = 10
    number_format = workbook.add_format({'num_format': '#,##0.00'})
    number_format.set_font_name(font_type)
    number_format.set_font_size(font_size)
    normal_cell_format = workbook.add_format()
    normal_cell_format.set_pattern(1)  # This is optional when using a solid fill.
    normal_cell_format.set_bg_color('#d7e4bc')
    normal_cell_format.set_border(1)
    normal_cell_format.set_font_name(font_type)
    normal_cell_format.set_font_size(font_size)

    header_cell_format = workbook.add_format()
    header_cell_format.set_pattern(1)  # This is optional when using a solid fill.
    header_cell_format.set_bg_color('#d7e4bc')
    header_cell_format.set_border(1)
    header_cell_format.set_bottom(2)
    header_cell_format.set_font_name(font_type)
    header_cell_format.set_font_size(font_size)

    return workbook, header_cell_format, normal_cell_format, number_format


def col_size(string):
    return len(str(string)) + 5


def check_cols_max_size(max_list, keys):
    for idx, val in enumerate(keys):
        key_size = col_size(val)
        if max_list[idx] < key_size:
            max_list[idx] = key_size

    return max_list


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
        # Rachunek zysków i start
        self.gain_and_loses_dict, self.gain_and_loses_title_tag = self.find_gains_and_loses()
        self.gain_and_loses_pd, self.gain_and_loses_border_list = self.generate_gains_and_loses_pd()
        # Zestawienie zmian w kapitale
        self.change_in_capital_dict = self.find_change_in_capital()
        self.change_in_capital_pd, self.change_in_capital_border_list = self.generate_change_in_capital_pd()
        # Rachunek Przepływów
        self.capital_flow_dict, self.capital_flow_title_tag = self.find_capital_flow()
        self.capital_flow_pd, self.capital_flow_border_list = self.generate_capital_flow_pd()
        logging.debug('test')

    def list_all_root_tags(self):
        for elem in self.root:
            logging.debug(elem.tag)

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
                finded = self.recurse_search(elem, tag)
                if finded is not None:
                    return finded

        return None

    def search(self, root, tag):
        for elem in root:
            if tag in elem.tag:
                return elem
            else:
                return self.r_search(elem, tag)
        return None

    def r_search(self, root, tag):
        if tag in root.tag:
            return root
        else:
            for elem in root:
                return self.r_search(elem, tag)

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

    def flatten_dict_with_numbers(self, d, main_tag):
        new_dict = {}
        for k, v in d[main_tag].items():
            if type(v) == dict:
                new_dict[k] = {}
                new_dict = self.go_deeper(v, "KwotaA", "KwotaB", k, new_dict)
            elif float(v) is not None:
                new_dict[k] = Decimal(v)  # float(v)
        return new_dict

    def go_deeper(self, d, tag1, tag2, main_tag, new_d):
        for k, v in d.items():
            if k == tag1 or k == tag2:
                new_d[main_tag][k] = Decimal(v)
            else:
                new_d[k] = {}
                self.go_deeper(v, tag1, tag2, k, new_d)

        return new_d

    def find_gains_and_loses(self):
        por_tag = self.recurse_search(self.root, 'RZiSPor')
        kalk_tag = self.recurse_search(self.root, 'RZiSKalk')

        if por_tag is not None:
            main_elem = por_tag
            main_tag = 'RZiSPor'
        elif kalk_tag is not None:
            main_elem = kalk_tag
            main_tag = 'RZiSKalk'

        gains_and_loses_dict_xml = {}
        gains_and_loses_dict_xml = self.xml_tag_to_dict(main_elem, gains_and_loses_dict_xml)

        gains_and_loses_dict = self.flatten_dict_with_numbers(gains_and_loses_dict_xml, main_tag)

        return {'RZiS': gains_and_loses_dict}, main_tag

    def find_change_in_capital(self):
        main_tag = 'ZestZmianWKapitale'
        main_elem = self.recurse_search(self.root, main_tag)

        change_in_capital_dict_xml = {}
        change_in_capital_dict_xml = self.xml_tag_to_dict(main_elem, change_in_capital_dict_xml)

        change_in_capital_dict = self.flatten_dict_with_numbers(change_in_capital_dict_xml, main_tag)

        return {main_tag: change_in_capital_dict}

    def find_capital_flow(self):
        main_tag = 'PrzeplywyPosr'
        flow_posr_tag = self.recurse_search(self.root, main_tag)

        capital_flow_dict_xml = {}
        capital_flow_dict_xml = self.xml_tag_to_dict(flow_posr_tag, capital_flow_dict_xml)

        capital_flow_dict = self.flatten_dict_with_numbers(capital_flow_dict_xml, main_tag)

        for k, v in capital_flow_dict.items():
            if not v:
                capital_flow_dict[k] = {self.cost_tag_1: None, self.cost_tag_2: None}

        return {'RachPrzeplywow': capital_flow_dict}, main_tag

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
            introduction = self.travel_dict_reccur(k, v, introduction, 'WprowadzenieDoSprawozdaniaFinansowego',
                                                   'WprowadzenieDoSprawozdaniaFinansowego', None, level)

        return introduction

    def get_equity_plus_liabilities(self):
        level = 0
        passives = {}
        for k, v in self.financial_balance_dict['Bilans']['Pasywa'].items():
            passives = self.travel_dict_reccur(k, v, passives, 'Pasywa', 'Pasywa', None, level)

        recombined_passives = self.join_two_subtags("KwotaA", "KwotaB", "Pasywa", passives)
        return recombined_passives

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

                    temp_dict[cost_type] = Decimal(cost)
                new_dict[main_tag][k] = new_list
            elif float(v) is not None:
                new_dict[main_tag]["Koszta"][k] = Decimal(v)
        return new_dict

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
        workbook, header_cell_format, normal_cell_format, number_format = excel_style_format(workbook)

        col_max_size = [0, 0]
        i = 0
        for index, row in self.introduction_pd.iterrows():
            if self.introduction_border_list[i] == 1:
                worksheet.write(i, 0, row['Col1'], header_cell_format)
                col_max_size = check_cols_max_size(col_max_size, [row['Col1'], ""])
            else:
                worksheet.write(i, 0, row['Col2'], normal_cell_format)
                worksheet.write(i, 1, row['Col3'], normal_cell_format)
                col_max_size = check_cols_max_size(col_max_size, [row['Col2'], row['Col3']])
            i += 1

        for idx, col_siz in enumerate(col_max_size):
            worksheet.set_column(idx, idx, col_siz)

        workbook.close()

    def generate_passives_excel(self, out_file_name):
        # self.generate_excel(out_file_name, self.passives_dict, base_tag="Pasywa")
        self.generate_excel_from_numbers_pd(self.passives_pd, self.passives_border_list, out_file_name)

    def generate_assets_excel(self, out_file_name):
        self.generate_excel_from_numbers_pd(self.assets_pd, self.assets_border_list, out_file_name)

    def generate_assets_pd(self):
        return self.generate_costs_pd(self.assets_dict, base_tag="Aktywa")

    def generate_passives_pd(self):
        return self.generate_costs_pd(self.passives_dict, base_tag="Pasywa")

    def generate_gains_and_loses_pd(self):
        return self.generate_pd_from_number_dict(self.gain_and_loses_dict, 'RZiS')

    def generate_change_in_capital_pd(self):
        return self.generate_pd_from_number_dict(self.change_in_capital_dict, 'ZestZmianWKapitale')

    def generate_capital_flow_pd(self):
        return self.generate_pd_from_number_dict(self.capital_flow_dict, 'RachPrzeplywow')

    def generate_pd_from_number_dict(self, d, main_key):
        border_this_list = []
        df = pd.DataFrame(
            columns=["key", self.xsd_dict[main_key], self.xsd_dict[self.cost_tag_1], self.xsd_dict[self.cost_tag_2]])

        row = 1
        for k, v in d[main_key].items():
            if re.match(r'^[A-Z]$', k):
                border_this_list.append(1)
            else:
                border_this_list.append(0)

            df.loc[row] = [k.replace('_', '.'), self.xsd_dict[k], v[self.cost_tag_1], v[self.cost_tag_2]]
            row += 1

        df.set_index("key", inplace=True)
        return df, border_this_list

    def generate_gains_and_loses_excel(self, out_file_name):
        self.generate_excel_from_numbers_pd(self.gain_and_loses_pd, self.gain_and_loses_border_list, out_file_name,
                                            title_row=self.gain_and_loses_title_tag)

    def generate_change_in_capital_excel(self, out_file_name):
        self.generate_excel_from_numbers_pd(self.change_in_capital_pd, self.change_in_capital_border_list,
                                            out_file_name)

    def generate_capital_flow_excel(self, out_file_name):
        self.generate_excel_from_numbers_pd(self.capital_flow_pd, self.capital_flow_border_list, out_file_name,
                                            title_row=self.capital_flow_title_tag)

    def generate_excel_from_numbers_pd(self, df, df_border_list, out_file_name, title_row=None):
        workbook = xlsxwriter.Workbook(out_file_name)
        worksheet = workbook.add_worksheet()
        workbook, header_cell_format, normal_cell_format, number_format = excel_style_format(workbook)

        cols = df.columns
        col_max_size = [0, 0, 0]

        for idx, col in enumerate(cols):
            worksheet.write(0, idx, col, header_cell_format)

        col_max_size = check_cols_max_size(col_max_size, cols)
        i = 1
        if title_row is not None:
            worksheet.write(1, 0, "", header_cell_format)
            worksheet.write(1, 1, "", header_cell_format)
            worksheet.write(1, 2, "", header_cell_format)
            worksheet.merge_range(f'A2:C2', self.xsd_dict[title_row], header_cell_format)
            i += 1

        for index, row in df.iterrows():
            if df_border_list[i - 2] == 1:
                cell_format = header_cell_format
            else:
                cell_format = normal_cell_format

            if row[cols[1]] is None or row[cols[2]] is None:
                worksheet.write(i, 0, "", cell_format)
                worksheet.write(i, 1, "", cell_format)
                worksheet.write(i, 2, "", cell_format)
                worksheet.merge_range(i, 0, i, 2, row[cols[0]], header_cell_format)
            else:
                prefix = f"{index}. " if len(index) > 0 else ""
                worksheet.write(i, 0, f"{prefix}{row[cols[0]]}", cell_format)
                worksheet.write_number(i, 1, row[cols[1]], number_format)
                worksheet.write_number(i, 2, row[cols[2]], number_format)
                col_max_size = check_cols_max_size(col_max_size, [row[cols[0]], row[cols[1]], row[cols[2]]])
            i += 1

        for idx, col_siz in enumerate(col_max_size):
            worksheet.set_column(idx, idx, col_siz)

        workbook.close()

    def generate_costs_pd(self, main_dict, base_tag="Aktywa", column_tag="Koszta"):
        border_this_list = []
        temp_dict = main_dict[base_tag].copy()
        df = pd.DataFrame(columns=["key", base_tag, self.xsd_dict[self.cost_tag_1], self.xsd_dict[self.cost_tag_2]])

        row = 0
        df.loc[row] = ["", self.xsd_dict[base_tag], temp_dict[column_tag][self.cost_tag_1],
                       temp_dict[column_tag][self.cost_tag_2]]
        border_this_list.append(0)

        row += 1
        del temp_dict['Koszta']

        border_count = 0
        for k, v in temp_dict.items():
            if type(v) == list:
                prefix = k.replace(f"{base_tag}_", "")
                df.loc[row] = [prefix, self.xsd_dict[k], v[0]['Koszta'][self.cost_tag_1],
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
                        df.loc[row] = [prefix, self.xsd_dict[list_k], list_v[self.cost_tag_1],
                                       list_v[self.cost_tag_2]]
                        row += 1
                        border_count += 1
        df.set_index("key", inplace=True)
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
        col_max_size = check_cols_max_size(col_max_size,
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

                        col_max_size = check_cols_max_size(col_max_size,
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
        files_dict = self.embedded_files_dict['DodatkoweInformacjeIObjasnieniaJednostkaInna'][
            'DodatkoweInformacjeIObjasnienia']
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


class FinancialXMLRulesCheck:
    def __init__(self, f_xml: FinancialXMLParser):
        self.f_xml = f_xml

        self.assets_colums = f_xml.assets_pd.columns
        self.passives_colums = f_xml.passives_pd.columns
        # "A.III": {"rows": [[]], "cols": [[self.assets_colums] * 5]}
        self.assets = {
            "": {"rows": [[self.f_xml.passives_pd.loc['']],
                          [self.f_xml.assets_pd.loc['A'], self.f_xml.assets_pd.loc['B'], self.f_xml.assets_pd.loc['C'],
                           self.f_xml.assets_pd.loc['D']]], "cols": [[self.passives_colums], [self.assets_colums] * 4]},
            "A": {"rows": [
                [self.f_xml.assets_pd.loc['A.I'], self.f_xml.assets_pd.loc['A.II'], self.f_xml.assets_pd.loc['A.III'],
                 self.f_xml.assets_pd.loc['A.IV'], self.f_xml.assets_pd.loc['A.V']]],
                "cols": [[self.assets_colums] * 5]},
            "A.I": {"rows": [[self.f_xml.assets_pd.loc['A.I.1'], self.f_xml.assets_pd.loc['A.I.2'],
                              self.f_xml.assets_pd.loc['A.I.3'],
                              self.f_xml.assets_pd.loc['A.I.4']]], "cols": [[self.assets_colums] * 4]},
            "A.II": {
                "rows": [[self.f_xml.assets_pd.loc['A.II.1'], self.f_xml.assets_pd.loc['A.II.2'],
                          self.f_xml.assets_pd.loc['A.II.3']]],
                "cols": [[self.assets_colums] * 3]},
            "A.II.1": {"rows": [[self.f_xml.assets_pd.loc['A.II.1.A'], self.f_xml.assets_pd.loc['A.II.1.B'],
                                 self.f_xml.assets_pd.loc['A.II.1.C'], self.f_xml.assets_pd.loc['A.II.1.D'],
                                 self.f_xml.assets_pd.loc['A.II.1.E']]], "cols": [[self.assets_colums] * 5]},
            "A.III": {"rows": [[self.f_xml.assets_pd.loc['A.III.1'], self.f_xml.assets_pd.loc['A.III.2'],
                                self.f_xml.assets_pd.loc['A.III.3']]], "cols": [[self.assets_colums] * 3]},
            "A.IV": {"rows": [[self.f_xml.assets_pd.loc['A.IV.1'], self.f_xml.assets_pd.loc['A.IV.2'],
                               self.f_xml.assets_pd.loc['A.IV.3'], self.f_xml.assets_pd.loc['A.IV.4']]],
                     "cols": [[self.assets_colums] * 4]},
            "A.IV.3": {"rows": [[self.f_xml.assets_pd.loc['A.IV.3.A'], self.f_xml.assets_pd.loc['A.IV.3.B'],
                                 self.f_xml.assets_pd.loc['A.IV.3.C']]], "cols": [[self.assets_colums] * 3]},
            "A.IV.3.A": {"rows": [[self.f_xml.assets_pd.loc['A.IV.3.A.1'], self.f_xml.assets_pd.loc['A.IV.3.A.2'],
                                   self.f_xml.assets_pd.loc['A.IV.3.A.3'], self.f_xml.assets_pd.loc['A.IV.3.A.4']]],
                         "cols": [[self.assets_colums] * 4]},
            "A.IV.3.B": {"rows": [[self.f_xml.assets_pd.loc['A.IV.3.B.1'], self.f_xml.assets_pd.loc['A.IV.3.B.2'],
                                   self.f_xml.assets_pd.loc['A.IV.3.B.3'], self.f_xml.assets_pd.loc['A.IV.3.B.4']]],
                         "cols": [[self.assets_colums] * 4]},
            "A.IV.3.C": {"rows": [[self.f_xml.assets_pd.loc['A.IV.3.C.1'], self.f_xml.assets_pd.loc['A.IV.3.C.2'],
                                   self.f_xml.assets_pd.loc['A.IV.3.C.3'], self.f_xml.assets_pd.loc['A.IV.3.C.4']]],
                         "cols": [[self.assets_colums] * 4]},
            "A.V": {"rows": [[self.f_xml.assets_pd.loc['A.V.1'], self.f_xml.assets_pd.loc['A.V.2']]],
                    "cols": [[self.assets_colums] * 2]},
            "B": {"rows": [
                [self.f_xml.assets_pd.loc['B.I'], self.f_xml.assets_pd.loc['B.II'], self.f_xml.assets_pd.loc['B.III'],
                 self.f_xml.assets_pd.loc['B.IV']]], "cols": [[self.assets_colums] * 4]},
            "B.I": {"rows": [
                [self.f_xml.assets_pd.loc['B.I.1'], self.f_xml.assets_pd.loc['B.I.2'],
                 self.f_xml.assets_pd.loc['B.I.3'],
                 self.f_xml.assets_pd.loc['B.I.4'], self.f_xml.assets_pd.loc['B.I.5']]],
                "cols": [[self.assets_colums] * 5]},
            "B.II": {"rows": [[self.f_xml.assets_pd.loc['B.II.1'], self.f_xml.assets_pd.loc['B.II.2'],
                               self.f_xml.assets_pd.loc['B.II.3']]], "cols": [[self.assets_colums] * 3]},
            "B.II.1": {"rows": [[self.f_xml.assets_pd.loc['B.II.1.A'], self.f_xml.assets_pd.loc['B.II.1.B']]],
                       "cols": [[self.assets_colums] * 2]},
            "B.II.1.A": {"rows": [[self.f_xml.assets_pd.loc['B.II.1.A.1'], self.f_xml.assets_pd.loc['B.II.1.A.2']]],
                         "cols": [[self.assets_colums] * 2]},
            "B.II.2": {"rows": [[self.f_xml.assets_pd.loc['B.II.2.A'], self.f_xml.assets_pd.loc['B.II.2.B']]],
                       "cols": [[self.assets_colums] * 2]},
            "B.II.2.A": {"rows": [[self.f_xml.assets_pd.loc['B.II.2.A.1'], self.f_xml.assets_pd.loc['B.II.2.A.2']]],
                         "cols": [[self.assets_colums] * 2]},
            "B.II.3": {"rows": [[self.f_xml.assets_pd.loc['B.II.3.A'], self.f_xml.assets_pd.loc['B.II.3.B'],
                                 self.f_xml.assets_pd.loc['B.II.3.C'], self.f_xml.assets_pd.loc['B.II.3.D']]],
                       "cols": [[self.assets_colums] * 4]},
            "B.II.3.A": {"rows": [[self.f_xml.assets_pd.loc['B.II.3.A.1'], self.f_xml.assets_pd.loc['B.II.3.A.2']]],
                         "cols": [[self.assets_colums] * 2]},
            "B.III": {"rows": [[self.f_xml.assets_pd.loc['B.III.1'], self.f_xml.assets_pd.loc['B.III.2']]],
                      "cols": [[self.assets_colums] * 2]},
            "B.III.1": {"rows": [[self.f_xml.assets_pd.loc['B.III.1.A'], self.f_xml.assets_pd.loc['B.III.1.B'],
                                  self.f_xml.assets_pd.loc['B.III.1.C']]],
                        "cols": [[self.assets_colums] * 3]},
            "B.III.1.A": {"rows": [[self.f_xml.assets_pd.loc['B.III.1.A.1'], self.f_xml.assets_pd.loc['B.III.1.A.2'],
                                    self.f_xml.assets_pd.loc['B.III.1.A.3'], self.f_xml.assets_pd.loc['B.III.1.A.4']]],
                          "cols": [[self.assets_colums] * 4]},
            "B.III.1.B": {"rows": [[self.f_xml.assets_pd.loc['B.III.1.B.1'], self.f_xml.assets_pd.loc['B.III.1.B.2'],
                                    self.f_xml.assets_pd.loc['B.III.1.B.3'], self.f_xml.assets_pd.loc['B.III.1.B.4']]],
                          "cols": [[self.assets_colums] * 4]},
            "B.III.1.C": {"rows": [[self.f_xml.assets_pd.loc['B.III.1.C.1'], self.f_xml.assets_pd.loc['B.III.1.C.2'],
                                    self.f_xml.assets_pd.loc['B.III.1.C.3']]],
                          "cols": [[self.assets_colums] * 3]}
        }

        # self.assets = {
        #     "": [{1: f_xml.assets_pd.loc[''][self.assets_colums[1]] - f_xml.passives_pd.loc[''][self.passives_colums[1]],
        #           2: f_xml.assets_pd.loc[''][self.assets_colums[2]] - f_xml.passives_pd.loc[''][self.passives_colums[2]]},
        #          {1: f_xml.assets_pd.loc[''][self.assets_colums[1]] - (
        #                  f_xml.assets_pd.loc['A'][self.assets_colums[1]] + f_xml.assets_pd.loc['B'][self.assets_colums[1]] +
        #                  f_xml.assets_pd.loc['C'][self.assets_colums[1]] + f_xml.assets_pd.loc['D'][self.assets_colums[1]]),
        #           2: f_xml.assets_pd.loc[''][self.assets_colums[2]] - (
        #                   f_xml.assets_pd.loc['A'][self.assets_colums[2]] + f_xml.assets_pd.loc['B'][self.assets_colums[2]] +
        #                   f_xml.assets_pd.loc['C'][self.assets_colums[2]] + f_xml.assets_pd.loc['D'][self.assets_colums[2]]),
        #           }
        #          ],
        #     "A": [{1: f_xml.assets_pd.loc['A'][self.assets_colums[1]] - (
        #             f_xml.assets_pd.loc['A.I'][self.assets_colums[1]] + f_xml.assets_pd.loc['A.II'][self.assets_colums[1]] +
        #             f_xml.assets_pd.loc['A.III'][self.assets_colums[1]] + f_xml.assets_pd.loc['A.IV'][self.assets_colums[1]] +
        #             f_xml.assets_pd.loc['A.V'][self.assets_colums[1]]),
        #            2: f_xml.assets_pd.loc['A'][self.assets_colums[2]] - (
        #                    f_xml.assets_pd.loc['A.I'][self.assets_colums[2]] + f_xml.assets_pd.loc['A.II'][
        #                self.assets_colums[2]] + f_xml.assets_pd.loc['A.III'][self.assets_colums[2]] +
        #                    f_xml.assets_pd.loc['A.IV'][self.assets_colums[2]] + f_xml.assets_pd.loc['A.V'][
        #                        self.assets_colums[2]])
        #
        #            }],
        #     "A.I": [{1: f_xml.assets_pd.loc['A.I'][self.assets_colums[1]] - (
        #             f_xml.assets_pd.loc['A.I.1'][self.assets_colums[1]] + f_xml.assets_pd.loc['A.I.2'][self.assets_colums[1]] +
        #             f_xml.assets_pd.loc['A.I.3'][self.assets_colums[1]] + f_xml.assets_pd.loc['A.I.4'][
        #                 self.assets_colums[1]]),
        #              2: f_xml.assets_pd.loc['A.I'][self.assets_colums[2]] - (
        #                      f_xml.assets_pd.loc['A.I.1'][self.assets_colums[2]] + f_xml.assets_pd.loc['A.I.2'][
        #                  self.assets_colums[2]] + f_xml.assets_pd.loc['A.I.3'][self.assets_colums[2]] +
        #                      f_xml.assets_pd.loc['A.I.4'][self.assets_colums[2]])}
        #
        #             ],
        #     "A.II": [{1: f_xml.assets_pd.loc['A.II'][self.assets_colums[1]] - (
        #             f_xml.assets_pd.loc['A.II.1'][self.assets_colums[1]] + f_xml.assets_pd.loc['A.II.2'][
        #         self.assets_colums[1]] + f_xml.assets_pd.loc['A.II.3'][self.assets_colums[1]]),
        #               2: f_xml.assets_pd.loc['A.II'][self.assets_colums[2]] - (
        #                       f_xml.assets_pd.loc['A.II.1'][self.assets_colums[2]] + f_xml.assets_pd.loc['A.II.2'][
        #                   self.assets_colums[2]] + f_xml.assets_pd.loc['A.II.3'][self.assets_colums[2]])}]
        #
        # }

    def rule_check(self, index):
        index_rows = self.assets[index]['rows']
        index_cols = self.assets[index]['cols']

        res_list = []

        for i in range(len(index_rows)):
            sum_col1 = Decimal(0.0)
            sum_col2 = Decimal(0.0)
            internal_index_rows = index_rows[i]
            internal_index_cols = index_cols[i]
            for j in range(len(index_rows[i])):
                sum_col1 += internal_index_rows[j][internal_index_cols[j][1]]
                sum_col2 += internal_index_rows[j][internal_index_cols[j][2]]
            sum_col1 = sum_col1 - self.f_xml.assets_pd.loc[index][self.assets_colums[1]]
            sum_col2 = sum_col2 - self.f_xml.assets_pd.loc[index][self.assets_colums[2]]
            bool_sub_col1 = sum_col1 == Decimal(0.0)
            bool_sub_col2 = sum_col2 == Decimal(0.0)
            res_list.append((bool_sub_col1, bool_sub_col2))

        return res_list

    def to_excel(self, out_file):
        workbook = xlsxwriter.Workbook(out_file)
        worksheet = workbook.add_worksheet()
        true_false_format = workbook.add_format({'num_format': '0;;"TRUE";"FALSE"'})
        workbook, header_cell_format, normal_cell_format, number_format = excel_style_format(workbook)

        df = self.f_xml.assets_pd
        df_border_list = self.f_xml.assets_border_list

        cols = df.columns
        col_max_size = [0, 0, 0]

        for idx, col in enumerate(cols):
            worksheet.write(0, idx, col, header_cell_format)

        col_max_size = check_cols_max_size(col_max_size, cols)
        i = 1
        for index, row in df.iterrows():
            if df_border_list[i - 2] == 1:
                cell_format = header_cell_format
            else:
                cell_format = normal_cell_format

            if row[cols[1]] is None or row[cols[2]] is None:
                worksheet.write(i, 0, "", cell_format)
                worksheet.write(i, 1, "", cell_format)
                worksheet.write(i, 2, "", cell_format)
                worksheet.merge_range(i, 0, i, 2, row[cols[0]], header_cell_format)
            else:
                prefix = f"{index}. " if len(index) > 0 else ""
                occurences = prefix.count('.')
                prefix = ' ' * occurences + prefix
                worksheet.write(i, 0, f"{prefix}{row[cols[0]]}", cell_format)
                worksheet.write_number(i, 1, row[cols[1]], number_format)
                worksheet.write_number(i, 2, row[cols[2]], number_format)
                col_max_size = check_cols_max_size(col_max_size, [row[cols[0]], row[cols[1]], row[cols[2]]])

                if index in self.assets:
                    sum_check = self.rule_check(index)
                    for idx, s_check in enumerate(sum_check):
                        # worksheet.write_number(i, 2 * idx + 3, s_check[0], number_format)
                        # worksheet.write_number(i, 2 * idx + 4, s_check[1], number_format)
                        worksheet.write_boolean(i, 2 * idx + 3, s_check[0],true_false_format)
                        worksheet.write_boolean(i, 2 * idx + 4, s_check[1],true_false_format)

                    # for idx, d in enumerate(self.assets[index]):
                    #     worksheet.write_number(i, 2 * idx + 3, d[1], number_format)
                    #     worksheet.write_number(i, 2 * idx + 4, d[2], number_format)

            i += 1

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
        fxml = FinancialXMLParser(path)
        # f_xml.generate_excel('Aktywa.xlsx')
        fxml.generate_assets_excel('Aktywa.xlsx')
        fxml.generate_passives_excel('Pasywa.xlsx')
        # passives = f_xml.get_equity_plus_liabilities()
    logging.debug(fxml.financial_balance_dict)
