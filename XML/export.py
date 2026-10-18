import hashlib
from dataclasses import dataclass, field
from pathlib import Path

from Settings import Settings
from XML.FinancialXMLParser import CheckResult, FinancialXMLParser, FinancialXMLRulesCheck


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
