import shutil
import zipfile
from decimal import Decimal
from pathlib import Path

import pandas as pd
import pytest

import cli
from Settings import Settings
from XML import FinancialXMLParser as parser_module
from XML.export import default_workbook_name, export_statement, export_summation, export_workbook
from XML.presentation import (CHANGE_COLUMN, CHANGE_PERCENT_COLUMN, Percent, hierarchy, short_year_headers,
                              with_year_change)
from XML.xsd_downloader import schema_cache_dir

DATA_DIR = Path(__file__).parent / "data"
SCHEMA_URL = "https://example.invalid/schemas/JednostkaInnaWZlotych.xsd"


@pytest.fixture
def xml_path(monkeypatch, tmp_path):
    monkeypatch.setattr(parser_module, "download_schema", lambda url: None)
    cache_dir = Path(schema_cache_dir(SCHEMA_URL))
    cache_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy(DATA_DIR / "statement.xsd", cache_dir / "statement.xsd")
    path = tmp_path / "statement.xml"
    shutil.copy(DATA_DIR / "statement.xml", path)
    return path


def test_hierarchy():
    levels = hierarchy(["", "A", "A.I", "A.I.1", "A.I.2", "A.II", "B", "B.I.1"])
    assert levels[""] == (None, 0, True)
    assert levels["A"] == ("", 1, True)
    assert levels["A.I.1"] == ("A.I", 3, False)
    assert levels["A.II"] == ("A", 2, False)
    # B.I is not in the table, so B.I.1 hangs directly under B
    assert levels["B.I.1"] == ("B", 2, False)
    assert levels["B"] == ("", 1, True)


def test_hierarchy_of_tax_info_keys():
    levels = hierarchy(["A", "A_RB", "A_RP", "B"])
    assert levels["A_RB"] == ("A", 1, False)
    assert levels["A"] == (None, 0, True)


def test_year_change():
    df = pd.DataFrame([["A", Decimal("150"), Decimal("100")], ["B", Decimal("5"), Decimal("0")],
                       ["C", None, Decimal("3")]], columns=["", "cur", "prev"], index=["A", "B", "C"], dtype=object)
    result = with_year_change(df)
    assert result.loc["A", CHANGE_COLUMN] == Decimal("50")
    assert result.loc["A", CHANGE_PERCENT_COLUMN] == Decimal("50")
    assert isinstance(result.loc["A", CHANGE_PERCENT_COLUMN], Percent)
    assert result.loc["B", CHANGE_COLUMN] == Decimal("5")
    assert result.loc["B", CHANGE_PERCENT_COLUMN] is None
    assert result.loc["C", CHANGE_COLUMN] is None
    # The original is left alone
    assert CHANGE_COLUMN not in df.columns


def test_short_year_headers():
    keywords = {"KwotaA": "bieżący", "KwotaB": "poprzedni"}
    assert short_year_headers(keywords, "2021-12-31") == {"bieżący": "Rok bieżący (2021)",
                                                          "poprzedni": "Rok poprzedni (2020)"}
    assert short_year_headers(keywords, "")["bieżący"] == "Rok bieżący"


def test_export_workbook(xml_path, tmp_path):
    settings = Settings()
    result = export_statement(xml_path, settings, tmp_path / "excel")
    summation = export_summation(result.parser, tmp_path / "excel", result.prefix)

    assert default_workbook_name(result) == "Testowa Spółka Sp. z o.o. 2022.xlsx"
    path = export_workbook(result, tmp_path / "out" / default_workbook_name(result), settings.get_keywords(),
                           summation)

    with zipfile.ZipFile(path) as workbook:
        sheets = workbook.read("xl/workbook.xml").decode("utf8")
        strings = workbook.read("xl/sharedStrings.xml").decode("utf8")
    for sheet in ["Wprowadzenie", "Aktywa", "Pasywa", "RZiS", "Podatek dochodowy", "Kontrola sum"]:
        assert f'name="{sheet}"' in sheets
    assert "Rok bieżący (2022)" in strings
    assert "Zmiana r/r" in strings
    assert "BŁĄD" in strings


def test_cli_workbook(xml_path, tmp_path, capsys):
    assert cli.main([str(xml_path), "-o", str(tmp_path), "--workbook"]) == 0
    assert (tmp_path / "Testowa Spółka Sp. z o.o. 2022.xlsx").exists()


def test_clipboard_text():
    from Widgets.StatementTable import clipboard_text

    assert clipboard_text(Decimal("1234.5"), ",") == "1234,50"
    assert clipboard_text(Percent("12.345"), ",") == "12,3%"
    assert clipboard_text(None, ",") == ""
    assert clipboard_text("a\tb", ",") == "a b"
