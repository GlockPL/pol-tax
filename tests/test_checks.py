import shutil
import zipfile
from decimal import Decimal
from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData, QUrl

from Settings import Settings
from XML import FinancialXMLParser as parser_module
from XML.FinancialXMLParser import CheckResult, FinancialXMLParser, FinancialXMLRulesCheck, describe_rule
from XML.export import export_summation
from XML.rules import RULES
from XML.xsd_downloader import schema_cache_dir

DATA_DIR = Path(__file__).parent / "data"
SCHEMA_URL = "https://example.invalid/schemas/JednostkaInnaWZlotych.xsd"


@pytest.fixture
def statement(monkeypatch, tmp_path):
    monkeypatch.setattr(parser_module, "download_schema", lambda url: None)
    cache_dir = Path(schema_cache_dir(SCHEMA_URL))
    cache_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy(DATA_DIR / "statement.xsd", cache_dir / "statement.xsd")
    xml_path = tmp_path / "statement.xml"
    shutil.copy(DATA_DIR / "statement.xml", xml_path)
    return FinancialXMLParser(str(xml_path), Settings())


def test_check_all_reports_values_and_differences(statement):
    results = FinancialXMLRulesCheck(statement).check_all()

    balance = [r for r in results if r.section == "assets" and r.index == "" and "Pasywa" in r.description]
    assert [(r.year, r.ok) for r in balance] == [(1, True), (2, True)]
    assert balance[0].actual == balance[0].expected == Decimal("150.00")

    # A has no sub-positions in the test file, so the range sum is 0 and the check fails with the full difference
    a_range = next(r for r in results if r.section == "assets" and r.index == "A" and r.year == 1)
    assert not a_range.ok
    assert a_range.difference == Decimal("100.00")
    assert a_range.description == "suma pozycji A.I do A.V"


def test_summation_excel_has_failures_sheet(statement, tmp_path):
    summation = export_summation(statement, tmp_path, "test")
    assert len(summation.failures) > 0
    with zipfile.ZipFile(summation.path) as workbook:
        sheets = workbook.read("xl/workbook.xml").decode("utf8")
    assert "Sumowanie" in sheets
    assert "Niezgodności" in sheets


def test_rule_descriptions():
    assert describe_rule({'range': 'A.II.1-3'}, 'A.II', 1) == "suma pozycji A.II.1 do A.II.3"
    assert describe_rule({'special': 'C+D-E'}, 'F', 1) == "C + D − E"
    assert describe_rule({'value': {'index': 'II', 'source': 'change_in_capital', 'year': 2}}, 'I', 1) == \
        "Zestawienie zmian w kapitale własnym II (rok poprzedni)"
    assert describe_rule({'value': {'index': '', 'source': 'passives'}}, '', 2) == "Pasywa razem (rok poprzedni)"
    operations = {'operations': [{'source': 'assets', 'index': 'B.I', 'year': 1},
                                 {'source': 'assets', 'index': 'B.I', 'year': 2}, {'operator': '-'},
                                 {'source': 'assets', 'index': 'A.V', 'year': 1},
                                 {'source': 'assets', 'index': 'A.V', 'year': 2}, {'operator': '-'},
                                 {'operator': '+'}]}
    assert describe_rule(operations, 'X', 1) == \
        "(Aktywa B.I (rok bieżący) − Aktywa B.I (rok poprzedni)) + (Aktywa A.V (rok bieżący) − Aktywa A.V (rok poprzedni))"


def test_short_term_liabilities_change_is_not_negated():
    # An increase of liabilities is a positive adjustment in the cash flow statement
    operations = RULES['capital_flow']['A.II.8'][0]['operations']
    assert {'operator': '*'} not in operations


def test_check_result_difference():
    result = CheckResult("assets", "A", "Aktywa trwałe", 1, Decimal("10"), Decimal("7"), "")
    assert result.difference == Decimal("3")
    assert not result.ok


def test_recent_files(tmp_path):
    settings = Settings()
    old = list(settings._recent_files)
    try:
        settings._recent_files = []
        files = []
        for i in range(10):
            path = tmp_path / f"f{i}.xml"
            path.write_text("<a/>", encoding="utf8")
            files.append(path)
            settings.add_recent_file(path)
        settings.add_recent_file(files[3])
        recent = settings.get_recent_files()
        assert recent[0] == str(files[3].resolve())
        assert len(recent) == Settings.MAX_RECENT_FILES
        assert len(set(recent)) == len(recent)

        files[9].unlink()
        assert str(files[9].resolve()) not in settings.get_recent_files()
    finally:
        settings._recent_files = old


def test_dropped_files_filtered_to_xml(tmp_path):
    from Widgets.MainWindow import xml_paths_from_mime

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(tmp_path / "a.xml")), QUrl.fromLocalFile(str(tmp_path / "b.pdf")),
                  QUrl("https://example.invalid/c.xml")])
    assert [p.name for p in xml_paths_from_mime(mime)] == ["a.xml"]
