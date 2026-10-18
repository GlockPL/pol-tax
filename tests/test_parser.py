import shutil
from decimal import Decimal
from pathlib import Path

import pytest

from Settings import Settings
from XML import FinancialXMLParser as parser_module
from XML.FinancialXMLParser import FinancialXMLParser, FinancialXMLRulesCheck
from XML.xsd_downloader import schema_cache_dir

DATA_DIR = Path(__file__).parent / "data"
SCHEMA_URL = "https://example.invalid/schemas/JednostkaInnaWZlotych.xsd"


@pytest.fixture
def statement(monkeypatch, tmp_path):
    # No network: serve the schema from the test data folder
    monkeypatch.setattr(parser_module, "download_schema", lambda url: None)
    cache_dir = Path(schema_cache_dir(SCHEMA_URL))
    cache_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy(DATA_DIR / "statement.xsd", cache_dir / "statement.xsd")
    xml_path = tmp_path / "statement.xml"
    shutil.copy(DATA_DIR / "statement.xml", xml_path)
    return FinancialXMLParser(str(xml_path), Settings())


def test_labels_come_only_from_own_schema(statement):
    # A different schema version in the cache must not override labels
    other_dir = Path(schema_cache_dir("https://example.invalid/schemas/other-version.xsd"))
    other_dir.mkdir(parents=True, exist_ok=True)
    (other_dir / "statement.xsd").write_text(
        '<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
        '<xsd:element name="Aktywa"><xsd:annotation><xsd:documentation>ZLA ETYKIETA</xsd:documentation>'
        '</xsd:annotation></xsd:element></xsd:schema>', encoding="utf8")

    assert statement.map_xsd()["Aktywa"] == "Aktywa razem"


def test_year_columns_are_current_then_previous(statement):
    cols = list(statement.assets_pd.columns)
    keywords = Settings().get_keywords()
    assert cols[1:] == [keywords["KwotaA"], keywords["KwotaB"]]
    assert statement.assets_pd.loc["A"][cols[1]] == Decimal("100.00")
    assert statement.assets_pd.loc["A"][cols[2]] == Decimal("60.00")


def test_rzis_detected(statement):
    assert statement.gains_and_losses_obj.main_key == "RZiSPor"
    assert statement.gain_and_loses_pd is not None


def test_embedded_file_cannot_escape_target_dir(statement, tmp_path):
    target = tmp_path / "out" / "doc"
    paths = statement.extract_embedded_files(target)
    assert paths == [target / "zalacznik.pdf"]
    assert paths[0].read_bytes().startswith(b"%PDF")


def test_missing_schema_location_is_reported(tmp_path):
    xml_path = tmp_path / "no_schema.xml"
    xml_path.write_text("<Root><Aktywa/></Root>", encoding="utf8")
    with pytest.raises(ValueError, match="schemaLocation"):
        FinancialXMLParser(str(xml_path), Settings())


def test_balance_rule_and_missing_references(statement):
    rules = FinancialXMLRulesCheck(statement)
    # Total assets equal total liabilities in both years
    assert rules.rule_check("", "assets")[0] == [True, True]
    # Cash flow rule references the equity statement, which this file doesn't have: no result instead of TRUE/FALSE
    rules.pds["capital_flow"]["pd"] = statement.assets_pd
    rules.rules["capital_flow"]["A"] = rules.rules["capital_flow"]["A.I"]
    assert rules.rule_check("A", "capital_flow") == [[None, None]]


def test_excel_export(statement, tmp_path):
    statement.generate_assets_excel(str(tmp_path / "aktywa.xlsx"))
    statement.generate_introduction_excel_from_pd(str(tmp_path / "wprowadzenie.xlsx"))
    FinancialXMLRulesCheck(statement).to_excel(str(tmp_path / "summation.xlsx"))
    assert (tmp_path / "aktywa.xlsx").stat().st_size > 0
    assert (tmp_path / "summation.xlsx").stat().st_size > 0


def test_empty_cells_stay_none(statement):
    # pandas 2+ turns None into NaN when rows are added one by one, which broke the summation Excel export
    df = statement.additional_info_pd
    header = df.loc["A"]
    assert all(header[col] is None for col in df.columns[1:])
    assert df.loc["B_RB"][df.columns[1]] == Decimal("10.00")
    assert df.loc["A_RB"][df.columns[1]] == Decimal("120.00")


def test_summation_with_tax_info_section(statement, tmp_path):
    FinancialXMLRulesCheck(statement).to_excel(str(tmp_path / "summation.xlsx"))
    statement.generate_additional_info_excel(str(tmp_path / "dodatkowe.xlsx"))
    assert (tmp_path / "summation.xlsx").stat().st_size > 0


def test_kalk_variant_drops_rzis_rules(statement):
    statement.gains_and_losses_obj.main_key = "RZiSKalk"
    rules = FinancialXMLRulesCheck(statement)
    assert "gain_and_loses" not in rules.rules
    assert "IA.6.A" not in rules.rules["change_in_capital"]
    assert "J" not in rules.rules.get("gain_and_loses", {})
