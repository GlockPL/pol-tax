import xml.etree.ElementTree as ET
import pandas as pd
from urllib.parse import unquote
from urllib.parse import urlparse
from pathlib import Path, PureWindowsPath
import xlsxwriter
import re
import copy
import base64
import logging
import string
import math

from decimal import Decimal, InvalidOperation

# My files Import
from XML.xsd_downloader import download_schema, schema_cache_dir
from XML.rules import RULES, SECTION_TITLES, YEAR_NAMES
from dataclasses import dataclass
from Settings import Settings

roman2arabic = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6, 'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10, 'XI': 11,
                'XII': 12, 'XIII': 13, 'XIV': 14, 'XV': 15, 'XVI': 16, 'XVII': 17, 'XVIII': 18, 'XIX': 19, 'XX': 20}
arabic2roman = {v: k for k, v in roman2arabic.items()}


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
    normal_cell_format.set_text_wrap()
    normal_cell_format.set_align('vcenter')

    header_cell_format = workbook.add_format()
    header_cell_format.set_pattern(1)  # This is optional when using a solid fill.
    header_cell_format.set_bg_color('#F8F8FF')
    header_cell_format.set_border(1)
    header_cell_format.set_bottom(2)
    header_cell_format.set_font_name(font_type)
    header_cell_format.set_font_size(font_size)
    header_cell_format.set_text_wrap()
    header_cell_format.set_align('vcenter')

    return workbook, header_cell_format, normal_cell_format, number_format


def strip_namespace(tag):
    last_char = tag.find('}')
    if last_char > 0:
        tag = tag[last_char + 1:]
    return tag


def xml_tag_to_dict(xml_tag, xml_dict):
    tag = strip_namespace(xml_tag.tag)

    text = xml_tag.text.strip() if xml_tag.text is not None else ""
    if text == "":
        xml_dict[tag] = {}
    else:
        xml_dict[tag] = text

    for elem in xml_tag:
        xml_dict[tag] = xml_tag_to_dict(elem, xml_dict[tag])

    return xml_dict


def object_dataframe(rows, columns, index=None):
    # dtype=object keeps None as None (pandas 2+ would turn empty cells into NaN floats) and Decimals as Decimals
    return pd.DataFrame(rows, columns=columns, index=index, dtype=object)


def is_missing(value):
    return value is None or (isinstance(value, float) and math.isnan(value))


def find_first_tag(root, tag_name):
    for element in root.iter():
        if element.tag.endswith(tag_name):
            return element
    return None


def col_size(input_string):
    return len(str(input_string)) - 2 # + 5


def check_cols_max_size(max_list, keys):
    for idx, val in enumerate(keys):
        key_size = col_size(val)
        if max_list[idx] < key_size:
            max_list[idx] = key_size

    return max_list


class FinancialXMLParts:
    cost_tag = "Kwota"
    special_tags = ['PozycjaUszczegolawiajaca']

    def __init__(self, main_keys, xml_root, xsd_dict):
        self.cost_tag_set = set()
        self.xsd_dict = xsd_dict
        self.xml_root = xml_root
        self.tag_element, self.main_key = self.determine_main_key_existance(main_keys)

        self.tree_dict = self.get_dict_from_xml()
        self.flatten_dict = self.flatten_xml_dict()
        self.pd, self.border_list = self.get_pd()

    def determine_main_key_existance(self, main_keys):
        for key in main_keys:
            tag_elem = find_first_tag(self.xml_root, key)
            if tag_elem is not None:
                return tag_elem, key
        return None, None

    def get_dict_from_xml(self):
        if self.tag_element is None:
            return None

        return xml_tag_to_dict(self.tag_element, {})

    def flatten_xml_dict(self):
        logging.debug(f"Flattening xml dict for {self.main_key}")
        if self.tree_dict is None:
            return None

        flatt_dict = {'Koszta': {}}
        for k, v in self.tree_dict[self.main_key].items():
            flatt_dict = self.recursive_cost_dict(k, v, None, flatt_dict)
        ret_dict = {self.main_key: flatt_dict['Koszta']}
        del flatt_dict['Koszta']
        return ret_dict | flatt_dict

    def get_pd(self):
        logging.debug(f"Generating pandas dataframe for {self.main_key}")
        if self.flatten_dict is None:
            return None, None

        border_this_list = []

        cost_tags = sorted(self.cost_tag_set)
        column_names = ["key", ''] + [self.xsd_dict.get(x, x) for x in cost_tags]
        rows = []

        for k, v in self.flatten_dict.items():
            if len(v) == 0:
                continue

            if re.match(r'^(.)\1*$', k.strip()):
                border_this_list.append(1)
            else:
                border_this_list.append(0)
            key_element = k.replace(f'{self.main_key}', '')
            key_element = key_element[1:] if len(key_element) > 0 and key_element[0] == "_" else key_element
            key_element = key_element.replace('_', '.')

            rows.append([key_element, self.xsd_dict.get(k, k)] + [v.get(x) for x in cost_tags])

        df = object_dataframe(rows, column_names)
        df.set_index("key", inplace=True)
        return df, border_this_list

    def get_excel(self, out_file_name):
        logging.debug(f"Generating excel file for {self.main_key}")

        workbook = xlsxwriter.Workbook(out_file_name)
        worksheet = workbook.add_worksheet()
        workbook, header_cell_format, normal_cell_format, number_format = excel_style_format(workbook)

        cols = self.pd.columns
        col_max_size = [0] * len(cols)

        for idx, col in enumerate(cols):
            worksheet.write(0, idx, col, header_cell_format)

        col_max_size = check_cols_max_size(col_max_size, cols)
        i = 1
        j = 0
        for index, row in self.pd.iterrows():
            if self.border_list[j] == 1:
                cell_format = header_cell_format
            else:
                cell_format = normal_cell_format

            if is_missing(row[cols[1]]) or is_missing(row[cols[2]]):
                worksheet.write(i, 0, "", cell_format)
                worksheet.write(i, 1, "", cell_format)
                worksheet.write(i, 2, "", cell_format)
                worksheet.merge_range(i, 0, i, 2, row[cols[0]], header_cell_format)
            else:
                prefix = f"{index}. " if len(index) > 0 else ""
                worksheet.write(i, 0, f"{prefix}{row[cols[0]]}", cell_format)
                size_check_list = [f"{prefix}{row[cols[0]]}"]
                for col_i in range(1, len(cols)):
                    worksheet.write_number(i, col_i, row[cols[col_i]], number_format)
                    size_check_list.append(row[cols[col_i]])
                col_max_size = check_cols_max_size(col_max_size, size_check_list)
            i += 1
            j += 1

        for idx, col_siz in enumerate(col_max_size):
            worksheet.set_column(idx, idx, width=col_siz)

        workbook.close()

    def recursive_cost_dict(self, key, value, up_key, main_dict):
        current_dict = main_dict['Koszta']
        if up_key is not None:
            current_dict = main_dict[up_key]

        if self.cost_tag in key:
            if key not in current_dict:
                self.cost_tag_set.add(key)
                current_dict[key] = Decimal(value)
        elif type(value) == dict:
            main_dict[key] = {}
            for k, v in value.items():
                self.recursive_cost_dict(k, v, key, main_dict)
        elif any(x in key for x in self.special_tags):
            detail_item_tag = f"{up_key}_{key}"
            self.xsd_dict[detail_item_tag] = value['NazwaPozycji']
            main_dict[detail_item_tag] = {}
            self.recursive_cost_dict(detail_item_tag, value['KwotyPozycji'], up_key, main_dict)

        return main_dict



class FinancialXMLParser:
    def __init__(self, xml_path: str, app_settings: Settings):
        self.app_settings = app_settings
        self.cost_tag_1 = "KwotaA"
        self.cost_tag_2 = "KwotaB"
        self.cost_tag_3 = "KwotaC"
        self.xml_path = xml_path
        self.tree = ET.parse(xml_path)
        self.root = self.tree.getroot()
        # Gets schema url from xml file
        self.schema_url = self.get_schema_url()
        if self.schema_url is None:
            raise ValueError(f"Plik {xml_path} nie zawiera adresu schematu (schemaLocation) - "
                             f"to nie jest e-sprawozdanie finansowe?")
        # Creates cache dir, separate for every schema
        self.cache_dir = schema_cache_dir(self.schema_url)
        Path(self.cache_dir).mkdir(parents=True, exist_ok=True)
        self.download_xsd()
        self.xsd_dict = self.map_xsd()
        self.xsd_dict_orig = self.xsd_dict.copy()
        self.xsd_dict.update(app_settings.get_keywords())
        # get introduction
        self.introduction_dict = self.find_introduction()
        self.introduction_pd, self.introduction_border_list = self.generate_introduction_pd()
        # Finacial objects setup
        self.assets_obj = FinancialXMLParts(["Aktywa"], self.root, self.xsd_dict)
        self.passives_obj = FinancialXMLParts(["Pasywa"], self.root, self.xsd_dict)
        self.gains_and_losses_obj = FinancialXMLParts(["RZiSPor", "RZiSKalk"], self.root, self.xsd_dict)
        self.change_in_capital_obj = FinancialXMLParts(["ZestZmianWKapitale"], self.root, self.xsd_dict)
        self.capital_flow_obj = FinancialXMLParts(["RachPrzeplywow"], self.root, self.xsd_dict)

        # Embedded file extraction
        self.list_of_dict_embedded_files = self.get_embedded_files()
        self.assets_pd, self.assets_border_list = self.assets_obj.pd, self.assets_obj.border_list
        self.passives_pd, self.passives_border_list = self.passives_obj.pd, self.passives_obj.border_list
        self.gain_and_loses_pd = self.gains_and_losses_obj.pd
        self.gain_and_loses_border_list = self.gains_and_losses_obj.border_list
        self.change_in_capital_pd = self.change_in_capital_obj.pd
        self.change_in_capital_border_list = self.change_in_capital_obj.border_list
        self.capital_flow_pd, self.capital_flow_border_list = self.capital_flow_obj.pd, self.capital_flow_obj.border_list
        # Informacja dodatkowa
        self.additional_info_dict = self.find_additional_info()
        self.additional_info_pd, self.additional_info_border_list = self.generate_additional_info_pd()

    def get_schema_url(self):
        schema_url = None
        for elem_k, elem_v in self.root.attrib.items():
            if strip_namespace(elem_k) == "schemaLocation":
                last_ocr = elem_v.rfind("http")
                schema_url = unquote(elem_v[last_ocr:])
                parse = urlparse(schema_url)
                if parse.netloc == 'www.mf.gov.pl':
                    schema_url = parse._replace(netloc='www.gov.pl')._replace(scheme='https').geturl()

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
                        val = re.sub(r'\[.*\]', '', val)
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

    def find_introduction(self):
        # WprowadzenieDoSprawozdaniaFinansowego
        introduction_tag_elem = find_first_tag(self.root, 'WprowadzenieDoSprawozdaniaFinansowego')
        if introduction_tag_elem is None:
            return {'WprowadzenieDoSprawozdaniaFinansowego': {}}
        introduction_dict_xml = xml_tag_to_dict(introduction_tag_elem, {})

        ret_dict = {'WprowadzenieDoSprawozdaniaFinansowego': {}}
        for k, v in introduction_dict_xml['WprowadzenieDoSprawozdaniaFinansowego'].items():
            match = re.findall('P_([0-9]{1,3})', k)
            if len(match) > 0 and int(match[0]) < 4:
                ret_dict['WprowadzenieDoSprawozdaniaFinansowego'][k] = v

        return ret_dict

    def find_additional_info(self):
        # InformacjaDodatkowaDotyczacaPodatkuDochodowego
        additional_info_tag_elem = find_first_tag(self.root, 'InformacjaDodatkowaDotyczacaPodatkuDochodowego')
        if additional_info_tag_elem is not None:
            additional_info_dict_xml = xml_tag_to_dict(additional_info_tag_elem, {})
            additional_info_dict = self.reorganize_additional_info_dict(
                additional_info_dict_xml['InformacjaDodatkowaDotyczacaPodatkuDochodowego'])
            return additional_info_dict

        return None

    def reorganize_additional_info_dict(self, xml_dict):
        new_dict = {}
        for k, v in xml_dict.items():
            if "Kwota" in v:
                new_v = self.convert_floats_to_decimal(v['Kwota'])
                new_dict[k] = new_v
            else:
                sub_dict = {}
                for kk, vv in v.items():
                    sub_dict[kk] = {'KwotaA': Decimal(vv), 'KwotaB': None, 'KwotaC': None}
                new_dict[k] = sub_dict

        return new_dict

    def convert_floats_to_decimal(self, data):
        if isinstance(data, dict):
            for key in data:
                data[key] = self.convert_floats_to_decimal(data[key])
        elif isinstance(data, list):
            for i in range(len(data)):
                data[i] = self.convert_floats_to_decimal(data[i])
        elif isinstance(data, str):
            try:
                data = Decimal(data)
            except InvalidOperation:
                pass
        return data

    def get_embedded_files(self):
        return [xml_tag_to_dict(element, {}) for element in self.root.iter()
                if element.tag.endswith('DodatkoweInformacjeIObjasnienia')]

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

    def company_info(self):
        """
        Basic facts about the company from the introduction, empty strings when not present
        """
        fields = self.flatten_dict(self.introduction_dict)

        def field(name):
            value = fields.get(name, "")
            return "" if value == "Header" else value

        return {"name": field("NazwaFirmy"), "city": field("Miejscowosc"), "nip": field("P_1D"),
                "krs": field("P_1E"), "date_from": field("DataOd"), "date_to": field("DataDo")}

    def generate_introduction_pd(self):
        rows = []
        border_this_list = []
        flatten_introduction = self.flatten_dict(self.introduction_dict)
        for k, v in flatten_introduction.items():
            try:
                row_name = self.xsd_dict[k]
            except KeyError:
                row_name = k

            if row_name == "":
                continue

            if v == "Header":
                rows.append([row_name, "", ""])
                border_this_list.append(1)
            else:
                rows.append(["", row_name, v])
                border_this_list.append(0)

        return object_dataframe(rows, ["Col1", "Col2", "Col3"]), border_this_list

    def generate_additional_info_pd(self):
        if self.additional_info_dict is not None:
            columns = [self.xsd_dict_orig["InformacjaDodatkowaDotyczacaPodatkuDochodowego"],
                       self.xsd_dict_orig[self.cost_tag_1], self.xsd_dict_orig[self.cost_tag_2],
                       self.xsd_dict_orig[self.cost_tag_3]]
            rows = []
            index = []
            border_list = []
            for k, v in self.additional_info_dict.items():
                k_index = self.xsd_dict_orig[k].split('.')[0].strip()
                index.append(k_index)
                rows.append([self.xsd_dict_orig[k], None, None, None])
                border_list.append(1)
                for kk, vv in v.items():
                    value_lst2 = [self.xsd_dict_orig[kk], None, None, None]
                    kk_index = kk.replace("_", ".")
                    for lp, (cost_k, cost_v) in enumerate(vv.items()):
                        value_lst2[lp + 1] = cost_v
                    index.append(f"{k_index}_{kk_index}")
                    rows.append(value_lst2)
                    border_list.append(0)

            return object_dataframe(rows, columns, index), border_list
        return None, None

    def generate_additional_info_excel(self, out_file_name):
        if self.additional_info_pd is None:
            return

        workbook = xlsxwriter.Workbook(out_file_name)
        worksheet = workbook.add_worksheet()
        workbook, header_cell_format, normal_cell_format, number_format = excel_style_format(workbook)
        cols = self.additional_info_pd.columns
        col_max_size = [0, 0, 0, 0]
        for idx, col in enumerate(cols):
            worksheet.write(0, idx, col, header_cell_format)
        col_max_size = check_cols_max_size(col_max_size, cols)
        i = 0
        for index, row in self.additional_info_pd.iterrows():
            if self.additional_info_border_list[i] == 1:
                worksheet.write(i + 1, 0, row[cols[0]], header_cell_format)
                col_max_size = check_cols_max_size(col_max_size, [row[cols[0]], "", "", ""])
            else:
                worksheet.write(i + 1, 0, row[cols[0]], normal_cell_format)
                worksheet.write(i + 1, 1, row[cols[1]], number_format)
                worksheet.write(i + 1, 2, row[cols[2]], number_format)
                worksheet.write(i + 1, 3, row[cols[3]], number_format)
                col_max_size = check_cols_max_size(col_max_size,
                                                   [row[cols[0]], row[cols[1]], row[cols[2]], row[cols[3]]])
            i += 1

        for idx, col_siz in enumerate(col_max_size):
            worksheet.set_column(idx, idx, col_siz)

        workbook.close()

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
        self.passives_obj.get_excel(out_file_name)

    def generate_assets_excel(self, out_file_name):
        self.assets_obj.get_excel(out_file_name)

    def generate_gains_and_loses_excel(self, out_file_name):
        self.gains_and_losses_obj.get_excel(out_file_name)

    def generate_change_in_capital_excel(self, out_file_name):
        self.change_in_capital_obj.get_excel(out_file_name)

    def generate_capital_flow_excel(self, out_file_name):
        self.capital_flow_obj.get_excel(out_file_name)

    def extract_embedded_files(self, target_dir: Path):
        files_list = []
        for tag_dict in self.list_of_dict_embedded_files:
            files_dict = tag_dict['DodatkoweInformacjeIObjasnienia']
            for k, v in files_dict.items():
                if k == "Plik":
                    files_list.append({'file_name': v['Nazwa'], 'content': v['Zawartosc']})

        target_dir = Path(target_dir)
        target_dir.mkdir(parents=True, exist_ok=True)
        file_paths = []
        for idx, elem in enumerate(files_list):
            # File names come from the XML, keep only the last path component so they cannot escape target_dir
            file_name = PureWindowsPath(elem['file_name']).name
            if file_name in ('', '.', '..'):
                file_name = f'zalacznik_{idx + 1}'
            concatenated_path = target_dir / file_name
            base64_bytes = elem['content'].encode('utf8')
            message_bytes = base64.b64decode(base64_bytes)

            with concatenated_path.open(mode='wb') as fh:
                fh.write(message_bytes)
            file_paths.append(concatenated_path)

        return file_paths


def decimal_to_bool(dec_val):
    return dec_val == 0.0


@dataclass
class CheckResult:
    section: str
    index: str
    label: str
    year: int
    actual: Decimal
    expected: Decimal
    description: str

    @property
    def difference(self):
        return self.actual - self.expected

    @property
    def ok(self):
        return self.difference == 0


def split_special(expression):
    """
    'C+D-E' -> (['C', 'D', 'E'], ['+', '-'])
    """
    indexes, operators, current = [], [], ''
    for char in expression:
        if char in '+-':
            indexes.append(current)
            operators.append(char)
            current = ''
        else:
            current += char
    indexes.append(current)
    return indexes, operators


def position_name(index):
    return index if index else "razem"


def describe_operand(source, index, year):
    if source == 'constants':
        return index
    return f"{SECTION_TITLES[source]} {position_name(index)} ({YEAR_NAMES[year]})"


def describe_rule(rule, index, year):
    """
    Human readable description of what the checked position should be equal to
    """
    def range_ends(range_str):
        elements = range_str.split('.')
        prefix = '.'.join(elements[:-1])
        first, last = elements[-1].split('-')
        return (f"{prefix}.{first}" if prefix else first), (f"{prefix}.{last}" if prefix else last)

    if 'range' in rule:
        first, last = range_ends(rule['range'])
        return f"suma pozycji {first} do {last}"
    if 'diff_range' in rule:
        first, last = range_ends(rule['diff_range'])
        return f"{first} pomniejszone o kolejne pozycje do {last}"
    if 'special' in rule:
        indexes, operators = split_special(rule['special'])
        text = indexes[0]
        for operator, operand in zip(operators, indexes[1:]):
            text += f" {'+' if operator == '+' else '−'} {operand}"
        return text
    if 'value' in rule:
        value = rule['value']
        if 'data' in value:
            return f"wartość {value['data']}"
        return describe_operand(value['source'], value['index'], value.get('year', year))
    if 'operations' in rule:
        stack = []
        for elem in rule['operations']:
            if 'index' in elem:
                stack.append(describe_operand(elem['source'], elem['index'], elem['year']))
            else:
                right, left = stack.pop(), stack.pop()
                operator = {'+': '+', '-': '−', '*': '×'}[elem['operator']]
                stack.append(f"({left} {operator} {right})")
        text = stack.pop()
        # The outermost operation is always fully wrapped in parentheses
        has_operator = any('operator' in elem for elem in rule['operations'])
        return text[1:-1] if has_operator else text
    return ""


class FinancialXMLRulesCheck:
    def __init__(self, f_xml: FinancialXMLParser):
        logging.debug(f"Beggining rulse check for {f_xml.xml_path}")
        self.f_xml = f_xml

        self.pds = {'assets': {'pd': self.f_xml.assets_pd, 'border_list': self.f_xml.assets_border_list},
                    'passives': {'pd': self.f_xml.passives_pd, 'border_list': self.f_xml.passives_border_list},
                    'gain_and_loses': {'pd': self.f_xml.gain_and_loses_pd,
                                       'border_list': self.f_xml.gain_and_loses_border_list},
                    'change_in_capital': {'pd': self.f_xml.change_in_capital_pd,
                                          'border_list': self.f_xml.change_in_capital_border_list},
                    'capital_flow': {'pd': self.f_xml.capital_flow_pd,
                                     'border_list': self.f_xml.capital_flow_border_list},
                    'additional_info': {'pd': self.f_xml.additional_info_pd,
                                        'border_list': self.f_xml.additional_info_border_list},
                    'constants': {'pd': self.constants_pd()}
                    }

        self.rules = copy.deepcopy(RULES)
        if self.f_xml.gains_and_losses_obj.main_key != 'RZiSPor':
            self.drop_gain_and_loses_rules()

    def drop_gain_and_loses_rules(self):
        """
        Rules were written for the comparative P&L variant (RZiSPor). For the calculation variant (RZiSKalk) the
        letters mean different positions, so every rule that uses the P&L would compare the wrong numbers.
        """
        logging.info("RZiS is not in the comparative variant, skipping rules that depend on it")
        self.rules.pop('gain_and_loses', None)
        for subset_rules in self.rules.values():
            for index in list(subset_rules.keys()):
                subset_rules[index] = [rule for rule in subset_rules[index]
                                       if rule.get('value', {}).get('source') != 'gain_and_loses'
                                       and not any(op.get('source') == 'gain_and_loses'
                                                   for op in rule.get('operations', []))]
                if len(subset_rules[index]) == 0:
                    del subset_rules[index]

    def constants_pd(self):
        indexes = ['-1', '0', '1']
        data = {'key': indexes,
                'Name': ['Minus_1', 'Zero', 'One'],
                'Year_1': [-1, 0, 1],
                'Year_2': [-1, 0, 1]}
        consts_df = pd.DataFrame(data)
        consts_df.set_index("key", inplace=True)
        return consts_df

    def letter_range_to_list(self, letter_range):
        all_letters = string.ascii_uppercase
        # Convert the range of letters to a list
        letters = [all_letters.index(l) for l in letter_range.replace('-', '')]
        letters = [all_letters[i] for i in range(letters[0], letters[-1] + 1)]
        return letters

    def number_range_to_list(self, number_range):
        split = number_range.split('-')
        number_list = [i for i in range(int(split[0]), int(split[1]) + 1)]
        return number_list

    def roman_numbers_range_to_list(self, roman_n_range):
        split = roman_n_range.split('-')
        number_list = [i for i in range(roman2arabic[split[0]], roman2arabic[split[1]] + 1)]
        roman_number_list = [arabic2roman[num] for num in number_list]
        return roman_number_list

    def range_type_to_list(self, range_str):
        if range_str[0] == 'A':
            return self.letter_range_to_list(range_str)
        elif range_str[0].isdigit():
            return self.number_range_to_list(range_str)
        elif range_str[0] in ['I', 'V', 'X', 'L', 'C', 'D', 'M']:
            return self.roman_numbers_range_to_list(range_str)
        else:
            return []

    def create_full_range_list(self, index_range):
        index_elements = index_range.split('.')
        range_str = index_elements[-1]
        range_list = self.range_type_to_list(range_str)
        if len(index_elements) > 1:
            index_sub_elements = index_elements[:len(index_elements) - 1]
            index_sub_elements = '.'.join(index_sub_elements)
            range_list = [f'{index_sub_elements}.{x}' for x in range_list]
        return range_list

    def reverse_polish_notation(self, expression_list):
        """
        Returns None when the statement referenced by the expression is missing from the file
        """
        stack = []
        for elem in expression_list:
            if 'index' in elem:
                df = self.pds[elem['source']]['pd']
                if df is None:
                    return None

                cols = df.columns
                val = self.get_pd_val(df, elem['index'], elem['year'], cols)
                stack.append(val)
            elif 'operator' in elem:
                right = stack.pop()
                left = stack.pop()

                if elem['operator'] == '+':
                    res = left + right
                elif elem['operator'] == '-':
                    res = left - right
                elif elem['operator'] == '*':
                    res = left * right
                stack.append(res)

        return stack.pop()

    def get_pd_val(self, df, index, year, columns, return_error=False):
        try:
            val = df.loc[index][columns[year]]
            error = False
        except (KeyError, IndexError):
            val = 0
            error = True

        if is_missing(val):
            val = 0

        if return_error:
            return val, error

        return val

    def evaluate_rule(self, index, subset_key):
        """
        For every rule of the position and every checked year returns (actual, expected) or None when the
        rule can't be checked because data it refers to is missing from the file
        """
        rule = self.rules[subset_key][index]
        res_list = []

        picked_pd = pd.DataFrame(self.pds[subset_key]['pd'])
        p_pd_columns = list(picked_pd.columns)

        for list_elem in rule:
            year_sum = []
            for year in list_elem['years']:
                # Wartość, którą teraz badamy
                checked_value = self.get_pd_val(picked_pd, index, year, p_pd_columns)
                # W tym przypadku bierze się zapisany w sposób zakresu literami, cyframi łacińskimi albo rzymksimi ciąg znaków
                # Np. A.I-VI i zamienia na listę indeksów, jeden po drugim i sumuje się wartości pod tymi indeksami
                if 'range' in list_elem:
                    index_range = list_elem['range']
                    range_list = self.create_full_range_list(index_range)
                    expected = Decimal(0)
                    for dec_val_ind in range_list:
                        expected += self.get_pd_val(picked_pd, dec_val_ind, year, p_pd_columns)
                elif 'value' in list_elem:
                    # Jeżeli pod indeksem value jest indeks data bierzę się gołą wartość i odejmuje od aktualnie
                    # sprawdzanej wartości, najgorsza możliwość bo nie jak przypilnować indeksu
                    if 'data' in list_elem['value']:
                        expected = list_elem['value']['data']
                    else:
                        # Bierze indeks, dataframe i odejmuje wyciągniętą wartość dla danego roku, jeżeli rok jest inny
                        # niż ten, który jest w indeksie szukania dla sprawdzanej wartości
                        picked_year = year
                        if 'year' in list_elem['value']:
                            picked_year = list_elem['value']['year']
                        value_idx = list_elem['value']['index']
                        value_src = self.pds[list_elem['value']['source']]['pd']
                        if value_src is None:
                            # The statement this rule compares with is not in the file
                            year_sum.append(None)
                            continue
                        value_src_cols = list(value_src.columns)
                        ref_val, missing = self.get_pd_val(value_src, value_idx, picked_year, value_src_cols,
                                                           return_error=True)
                        if missing:
                            year_sum.append(None)
                            continue
                        expected = ref_val
                elif 'diff_range' in list_elem:
                    # Podobnie jak w range, tylko tutaj jest ciąg liczb odejmowanych od siebie
                    index_range = list_elem['diff_range']
                    range_list = self.create_full_range_list(index_range)
                    expected = self.get_pd_val(picked_pd, range_list[0], year, p_pd_columns)
                    for dec_val_ind in range_list[1:]:
                        expected -= self.get_pd_val(picked_pd, dec_val_ind, year, p_pd_columns)
                elif 'special' in list_elem:
                    # Wyciąga indeksy i znaki operacji z ciągu znaków, który zawiera zapisane tekstem operacje, których
                    # nie da się zapisać w prosty sposób na przykłąd zakresem, bo wartości są losowe np. A.II+B.I.3
                    indexes, operators = split_special(list_elem['special'])
                    expected = self.get_pd_val(picked_pd, indexes[0], year, p_pd_columns)
                    for i in range(len(operators)):
                        if operators[i] == '+':
                            expected += self.get_pd_val(picked_pd, indexes[i + 1], year, p_pd_columns)
                        elif operators[i] == '-':
                            expected -= self.get_pd_val(picked_pd, indexes[i + 1], year, p_pd_columns)
                elif 'operations' in list_elem:
                    rpn = self.reverse_polish_notation(list_elem['operations'])
                    if rpn is None:
                        year_sum.append(None)
                        continue
                    expected = rpn
                year_sum.append((checked_value, expected))

            res_list.append(year_sum)

        return res_list

    def rule_check(self, index, subset_key):
        return [[None if pair is None else decimal_to_bool(pair[0] - pair[1]) for pair in year_pairs]
                for year_pairs in self.evaluate_rule(index, subset_key)]

    def checked_rows(self):
        """
        Yields (section key, index, row) of every position that has rules, the same rows the summation sheet checks
        """
        for rule_key in self.rules.keys():
            df = self.pds[rule_key]['pd']
            if df is None:
                continue
            cols = df.columns
            for index, row in df.iterrows():
                if is_missing(row[cols[1]]) or index not in self.rules[rule_key]:
                    continue
                yield rule_key, index, row

    def check_all(self):
        """
        Returns a CheckResult for every rule, position and year that could be checked
        """
        results = []
        for rule_key, index, row in self.checked_rows():
            label = row[row.index[0]]
            for rule, year_pairs in zip(self.rules[rule_key][index], self.evaluate_rule(index, rule_key)):
                for year, pair in zip(rule['years'], year_pairs):
                    if pair is None:
                        continue
                    actual, expected = pair
                    results.append(CheckResult(section=rule_key, index=index, label=label, year=year,
                                               actual=Decimal(actual), expected=Decimal(expected),
                                               description=describe_rule(rule, index, year)))
        return results

    def height_check(self, text, worksheet, index):
        char_line_limit = 75
        wrap_text_count = len(text)
        height = 20
        if wrap_text_count > char_line_limit:
            # If there is more than one line of wrapped text, adjust the row height
            times = wrap_text_count // char_line_limit
            height = height * (times + 1)
        # Set the row height
        worksheet.set_row(index, height)
        return worksheet

    def to_excel(self, out_file, results=None):
        if results is None:
            results = self.check_all()
        workbook = xlsxwriter.Workbook(out_file)
        workbook.formats[0].set_font_size(10)
        worksheet = workbook.add_worksheet("Sumowanie")
        workbook, header_cell_format, normal_cell_format, number_format = excel_style_format(workbook)
        # Sets first columns width
        worksheet.set_column(0, 0, 100)
        col_max_size = [0, 0, 0, 0]

        i = 0
        for rule_key in self.rules.keys():
            df = self.pds[rule_key]['pd']
            df_border_list = self.pds[rule_key]['border_list']
            # Sprawdza czy akurat ten plik xml nie zawiera tej podkategorii
            if df is None:
                continue

            cols = df.columns

            for idx, col in enumerate(cols):
                worksheet.write(i, idx, col, header_cell_format)
            worksheet = self.height_check(cols[0], worksheet, i)
            i += 1
            j = 0
            col_max_size = check_cols_max_size(col_max_size, cols[:3])
            for index, row in df.iterrows():
                if df_border_list[j] == 1:
                    cell_format = header_cell_format
                    height = 40
                else:
                    height = 20
                    cell_format = normal_cell_format

                if is_missing(row[cols[1]]):
                    worksheet.write(i, 0, "", cell_format)
                    worksheet.write(i, 1, "", cell_format)
                    worksheet.write(i, 2, "", cell_format)
                    text_to_merge = row[cols[0]]
                    worksheet.merge_range(i, 0, i, 2, text_to_merge, header_cell_format)
                else:
                    prefix = f"{index}. " if len(index) > 0 else ""
                    occurrences = prefix.count('.')
                    prefix = ' ' * occurrences + prefix
                    first_cell_text = f"{prefix}{row[cols[0]]}"
                    worksheet.write(i, 0, first_cell_text, cell_format)
                    for c in range(1, len(cols)):
                        cell_data = row[cols[c]]
                        if type(cell_data) == Decimal:
                            worksheet.write_number(i, c, cell_data, number_format)
                        elif not is_missing(cell_data):
                            worksheet.write(i, c, cell_data)
                    char_line_limit = 75
                    wrap_text_count = len(first_cell_text)
                    if wrap_text_count > char_line_limit:
                        height = 20
                        # If there is more than one line of wrapped text, adjust the row height
                        times = wrap_text_count // char_line_limit
                        height = height * (times + 1)
                    # Set the row height
                    worksheet.set_row(i, height)

                    col_list = [it[0] for it in row[cols].items()]
                    col_max_size = check_cols_max_size(col_max_size, col_list)

                    if index in self.rules[rule_key]:
                        sum_res = self.rule_check(index, rule_key)
                        for idx, s_val in enumerate(sum_res):
                            for s_id, s_check in enumerate(s_val):
                                if s_check is not None:
                                    worksheet.write_boolean(i, 2 * idx + len(cols) + s_id, s_check)
                i += 1
                j += 1

        for idx, col_siz in enumerate(col_max_size):
            if idx == 0:
                col_siz = col_siz // 2
            worksheet.set_column(idx, idx, col_siz)

        self.write_failures_sheet(workbook, header_cell_format, normal_cell_format, number_format, results)
        workbook.close()

    def write_failures_sheet(self, workbook, header_cell_format, normal_cell_format, number_format, results):
        worksheet = workbook.add_worksheet("Niezgodności")
        headers = ["Sekcja", "Pozycja", "Nazwa", "Rok", "Wartość", "Oczekiwana", "Różnica", "Reguła"]
        widths = [28, 10, 60, 14, 18, 18, 18, 70]
        for col, (header, width) in enumerate(zip(headers, widths)):
            worksheet.write(0, col, header, header_cell_format)
            worksheet.set_column(col, col, width)
        failures = [result for result in results if not result.ok]
        if len(failures) == 0:
            worksheet.write(1, 0, "Wszystkie reguły są spełnione", normal_cell_format)
        for row, result in enumerate(failures, start=1):
            worksheet.write(row, 0, SECTION_TITLES[result.section], normal_cell_format)
            worksheet.write(row, 1, result.index or "razem", normal_cell_format)
            worksheet.write(row, 2, result.label, normal_cell_format)
            worksheet.write(row, 3, YEAR_NAMES[result.year], normal_cell_format)
            worksheet.write_number(row, 4, result.actual, number_format)
            worksheet.write_number(row, 5, result.expected, number_format)
            worksheet.write_number(row, 6, result.difference, number_format)
            worksheet.write(row, 7, result.description, normal_cell_format)

