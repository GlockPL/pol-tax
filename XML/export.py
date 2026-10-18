import hashlib
import re
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

import xlsxwriter

from Settings import Settings
from XML.FinancialXMLParser import CheckResult, FinancialXMLParser, FinancialXMLRulesCheck
from XML.presentation import YEAR_SECTIONS, Percent, hierarchy, short_year_headers, with_year_change
from XML.rules import SECTION_TITLES, YEAR_NAMES


@dataclass(frozen=True)
class Section:
    key: str
    file_suffix: str
    pd_attr: str
    border_attr: str
    generate_method: str


SECTIONS = [
    Section("introduction", "Wprowadzenie", "introduction_pd", "introduction_border_list",
            "generate_introduction_excel_from_pd"),
    Section("assets", "Aktywa", "assets_pd", "assets_border_list", "generate_assets_excel"),
    Section("passives", "Pasywa", "passives_pd", "passives_border_list", "generate_passives_excel"),
    Section("gain_and_loses", "RZiS", "gain_and_loses_pd", "gain_and_loses_border_list",
            "generate_gains_and_loses_excel"),
    Section("change_in_capital", "ZestZmianWKapitale", "change_in_capital_pd", "change_in_capital_border_list",
            "generate_change_in_capital_excel"),
    Section("capital_flow", "RachPrzeplywow", "capital_flow_pd", "capital_flow_border_list",
            "generate_capital_flow_excel"),
    Section("additional_info", "DodatkoweInfo", "additional_info_pd", "additional_info_border_list",
            "generate_additional_info_excel"),
]


@dataclass
class ExportResult:
    parser: FinancialXMLParser
    xml_path: Path
    prefix: str
    # Section key -> (excel path, dataframe, border list); sections missing from the file are left out
    sections: dict = field(default_factory=dict)
    attachments: list = field(default_factory=list)


def output_prefix(xml_path: Path) -> str:
    # File name for readability plus a hash of the full path, so same-named XMLs from different folders don't collide
    sha = hashlib.sha256(str(Path(xml_path).resolve()).encode())
    return f"{Path(xml_path).stem}_{sha.hexdigest()[:8]}"


def export_statement(xml_path: Path, settings: Settings, excel_dir: Path = None) -> ExportResult:
    """
    Parses the statement, writes one Excel file per section and extracts embedded attachments
    """
    xml_path = Path(xml_path)
    excel_dir = Path(excel_dir or settings.get_excel_dir())
    excel_dir.mkdir(parents=True, exist_ok=True)

    parser = FinancialXMLParser(str(xml_path), settings)
    result = ExportResult(parser=parser, xml_path=xml_path, prefix=output_prefix(xml_path))

    for section in SECTIONS:
        dataframe = getattr(parser, section.pd_attr)
        if dataframe is None:
            continue
        excel_path = excel_dir / f"{result.prefix}_{section.file_suffix}.xlsx"
        getattr(parser, section.generate_method)(str(excel_path))
        result.sections[section.key] = (excel_path, dataframe, getattr(parser, section.border_attr))

    result.attachments = parser.extract_embedded_files(settings.get_extracted_dir() / result.prefix)
    return result


@dataclass
class SummationResult:
    path: Path
    results: list[CheckResult]

    @property
    def failures(self):
        return [result for result in self.results if not result.ok]


def export_summation(parser: FinancialXMLParser, excel_dir: Path, prefix: str) -> SummationResult:
    """
    Runs the consistency check and writes it to Excel (all positions plus a sheet with the failed checks)
    """
    excel_dir = Path(excel_dir)
    excel_dir.mkdir(parents=True, exist_ok=True)
    summation_path = excel_dir / f"{prefix}_summation.xlsx"
    rules_check = FinancialXMLRulesCheck(parser)
    results = rules_check.check_all()
    rules_check.to_excel(str(summation_path), results)
    return SummationResult(path=summation_path, results=results)


# Excel limits sheet names to 31 characters
SHEET_NAMES = {
    "introduction": "Wprowadzenie",
    "assets": "Aktywa",
    "passives": "Pasywa",
    "gain_and_loses": "RZiS",
    "change_in_capital": "Zmiany w kapitale",
    "capital_flow": "Przepływy pieniężne",
    "additional_info": "Podatek dochodowy",
}


def default_workbook_name(result: ExportResult) -> str:
    info = result.parser.company_info()
    name = re.sub(r'[<>:"/\\|?*]', "", info["name"]).strip() or result.xml_path.stem
    year = info["date_to"][:4]
    return f"{name} {year}.xlsx" if year.isdigit() else f"{name}.xlsx"


def export_workbook(result: ExportResult, path: Path, keywords: dict, summation: SummationResult = None) -> Path:
    """
    One workbook with a sheet per part of the statement (positions indented by level, year-over-year change)
    and, when the check was run, a sheet with all check results
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    info = result.parser.company_info()
    workbook = xlsxwriter.Workbook(str(path))
    font = {"font_name": "Arial", "font_size": 10}
    formats = {
        "title": workbook.add_format({**font, "bold": True, "font_size": 13}),
        "header": workbook.add_format({**font, "bold": True, "bg_color": "#F2F2F2", "bottom": 1,
                                       "text_wrap": True, "valign": "vcenter"}),
        "key": workbook.add_format({**font, "font_color": "#555555"}),
        "amount": workbook.add_format({**font, "num_format": "#,##0.00"}),
        "amount_bold": workbook.add_format({**font, "num_format": "#,##0.00", "bold": True}),
        "percent": workbook.add_format({**font, "num_format": "+0.0%;-0.0%;0.0%"}),
        "text": workbook.add_format(font),
        "ok": workbook.add_format({**font, "font_color": "#2E7D32", "bold": True}),
        "failed": workbook.add_format({**font, "font_color": "#B02A37", "bold": True}),
    }
    name_formats = {}

    def name_format(depth, bold):
        if (depth, bold) not in name_formats:
            name_formats[(depth, bold)] = workbook.add_format({**font, "indent": min(depth, 15), "bold": bold})
        return name_formats[(depth, bold)]

    title = info["name"] or result.xml_path.stem
    details = "   ".join(part for part in [f"NIP {info['nip']}" if info["nip"] else "",
                                          f"KRS {info['krs']}" if info["krs"] else "",
                                          f"okres {info['date_from']} – {info['date_to']}"
                                          if info["date_from"] else ""] if part)

    for section in SECTIONS:
        if section.key not in result.sections:
            continue
        _, dataframe, _ = result.sections[section.key]
        sheet = workbook.add_worksheet(SHEET_NAMES[section.key])
        sheet.write(0, 0, title, formats["title"])
        sheet.write(1, 0, f"{SECTION_TITLES[section.key]}   {details}", formats["text"])

        if section.key == "introduction":
            sheet.set_column(0, 0, 45)
            sheet.set_column(1, 1, 35)
            sheet.set_column(2, 2, 50)
            for row, (_, values) in enumerate(dataframe.iterrows(), start=3):
                group, field_name, value = list(values)
                if group:
                    sheet.write(row, 0, group, name_format(0, True))
                else:
                    sheet.write(row, 1, field_name, formats["text"])
                    sheet.write(row, 2, value, formats["text"])
            continue

        if section.key in YEAR_SECTIONS:
            dataframe = with_year_change(dataframe.rename(
                columns=short_year_headers(keywords, info["date_to"])))
        columns = list(dataframe.columns)
        headers = ["Pozycja", columns[0] or "Nazwa"] + columns[1:]
        header_row = 3
        for col, header in enumerate(headers):
            sheet.write(header_row, col, header, formats["header"])
        sheet.set_row(header_row, 30)
        sheet.set_column(0, 0, 12)
        sheet.set_column(1, 1, 70)
        sheet.set_column(2, len(headers) - 1, 18)

        levels = hierarchy(dataframe.index)
        for row, (key, values) in enumerate(dataframe.iterrows(), start=header_row + 1):
            _, depth, has_children = levels[str(key)]
            sheet.write(row, 0, str(key) or "razem", formats["key"])
            sheet.write(row, 1, values.iloc[0], name_format(depth, has_children))
            for col, value in enumerate(values.iloc[1:], start=2):
                if isinstance(value, Percent):
                    sheet.write_number(row, col, value / 100, formats["percent"])
                elif isinstance(value, Decimal):
                    sheet.write_number(row, col, value, formats["amount_bold" if has_children else "amount"])
        sheet.freeze_panes(header_row + 1, 2)
        sheet.autofilter(header_row, 0, header_row + len(dataframe), len(headers) - 1)

    if summation is not None:
        sheet = workbook.add_worksheet("Kontrola sum")
        sheet.write(0, 0, title, formats["title"])
        sheet.write(1, 0, f"Kontrola sum: niezgodności {len(summation.failures)}, "
                          f"sprawdzonych reguł {len(summation.results)}", formats["text"])
        headers = ["Wynik", "Część", "Pozycja", "Nazwa", "Rok", "Wartość", "Oczekiwana", "Różnica", "Reguła"]
        widths = [8, 30, 10, 50, 14, 18, 18, 16, 80]
        for col, (header, width) in enumerate(zip(headers, widths)):
            sheet.write(3, col, header, formats["header"])
            sheet.set_column(col, col, width)
        # Failed checks first
        ordered = summation.failures + [r for r in summation.results if r.ok]
        for row, check in enumerate(ordered, start=4):
            sheet.write(row, 0, "OK" if check.ok else "BŁĄD", formats["ok" if check.ok else "failed"])
            sheet.write(row, 1, SECTION_TITLES[check.section], formats["text"])
            sheet.write(row, 2, check.index or "razem", formats["text"])
            sheet.write(row, 3, check.label, formats["text"])
            sheet.write(row, 4, YEAR_NAMES[check.year], formats["text"])
            sheet.write_number(row, 5, check.actual, formats["amount"])
            sheet.write_number(row, 6, check.expected, formats["amount"])
            sheet.write_number(row, 7, check.difference, formats["amount"])
            sheet.write(row, 8, check.description, formats["text"])
        sheet.freeze_panes(4, 0)
        sheet.autofilter(3, 0, 3 + len(ordered), len(headers) - 1)

    workbook.close()
    return path
