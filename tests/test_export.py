import json
import pickle
import tomllib
import shutil
from pathlib import Path

import pytest

import cli
from Other.Version import app_version, windows_version_tuple
from Settings import Settings
from XML import FinancialXMLParser as parser_module
from XML.export import export_statement, export_summation, output_prefix
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


def test_export_statement_writes_present_sections_only(xml_path, tmp_path):
    out_dir = tmp_path / "excel"
    result = export_statement(xml_path, Settings(), out_dir)

    assert set(result.sections) == {"introduction", "assets", "passives", "gain_and_loses", "additional_info"}
    for excel_path, _, _ in result.sections.values():
        assert excel_path.parent == out_dir
        assert excel_path.name.startswith(result.prefix)
        assert excel_path.exists()
    assert [p.name for p in result.attachments] == ["zalacznik.pdf"]

    summation = export_summation(result.parser, out_dir, result.prefix)
    assert summation.exists()


def test_output_prefix_differs_for_same_name_in_other_folder(tmp_path):
    a = output_prefix(tmp_path / "a" / "sprawozdanie.xml")
    b = output_prefix(tmp_path / "b" / "sprawozdanie.xml")
    assert a != b
    assert a.startswith("sprawozdanie_")


def test_cli_converts_and_reports_errors(xml_path, tmp_path, capsys):
    broken = tmp_path / "broken.xml"
    broken.write_text("<Root/>", encoding="utf8")

    exit_code = cli.main([str(xml_path), str(broken), "-o", str(tmp_path / "cli"), "--check"])

    out, err = capsys.readouterr()
    assert exit_code == 1
    assert "summation.xlsx" in out
    assert "broken.xml: BŁĄD" in err


def test_settings_saved_as_json():
    settings = Settings()
    settings.cleanup()
    saved = json.loads((settings.get_docs_path() / "settings.json").read_text(encoding="utf8"))
    assert saved["keywords"]["KwotaA"] == settings.get_keywords()["KwotaA"]
    assert Path(saved["excel_dir"]) == Path(settings.get_excel_dir())


def test_old_pickle_settings_are_migrated(tmp_path):
    settings = Settings()
    old_docs, old_settings_path = settings._docs_path, settings._settings_path
    try:
        settings._docs_path = tmp_path
        settings._settings_path = tmp_path / "settings.json"
        with (tmp_path / "settings.pickle").open("wb") as f:
            pickle.dump({"_excel_dir": tmp_path / "moje_excele", "_additional_keywords": {"KwotaA": "Bieżący"}}, f)

        settings.migrate_pickle()

        assert settings.get_keywords() == {"KwotaA": "Bieżący"}
        assert Path(settings.get_excel_dir()) == tmp_path / "moje_excele"
        assert not (tmp_path / "settings.pickle").exists()
        assert (tmp_path / "settings.json").exists()
    finally:
        settings._docs_path, settings._settings_path = old_docs, old_settings_path
        settings._additional_keywords = dict(Settings._additional_keywords)


def test_version_comes_from_pyproject():
    with (Path(__file__).parent.parent / "pyproject.toml").open("rb") as f:
        assert app_version() == tomllib.load(f)["project"]["version"]


def test_windows_version_tuple():
    assert windows_version_tuple("1.2.3", 45) == (1, 2, 3, 45)
    assert windows_version_tuple("1.2") == (1, 2, 0, 0)
