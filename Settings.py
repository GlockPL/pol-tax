import logging

from PySide6.QtCore import QStandardPaths

import pickle
import os
from pathlib import Path
import atexit


class Settings:
    _excel_dir = Path("./Excel")
    _extracted_files_dir = Path("./Extracted")
    # TODO: Paradox of opening setting file that can change it's location after closing
    _settings_path = Path("./settings.pickle")
    _docs_path = None
    _additional_keywords = {
        "KwotaA": "Kwota na dzień kończący bieżący rok obrotowy",
        "KwotaB": "Kwota na dzień kończący poprzedni rok obrotowy",
        "DataOd": "Okres od",
        "DataDo": "Okres do",
        "NazwaFirmy": "Nazwa Firmy",
        "Siedziba": "Siedziba",
        "P_1B": "",
        "P_1D": "NIP",
        "P_1E": "KRS",
        "KodPKD": "Kod PKD"
    }
    __instance = None

    def __new__(cls):
        if cls.__instance is None:
            cls.__instance = super(Settings, cls).__new__(cls)
            cls.__instance.__initialized = False
        return cls.__instance

    def __init__(self):
        if self.__initialized:
            return
        self.__initialized = True
        self._docs_path = self.find_docs_path()

        self._excel_dir = self._docs_path.joinpath(self._excel_dir)
        self._extracted_files_dir = self._docs_path.joinpath(self._extracted_files_dir)
        self._settings_path = self._docs_path.joinpath(self._settings_path)

        if self._settings_path.exists():
            with self._settings_path.open(mode='rb') as settings_file:
                settings_dict = pickle.load(settings_file)
                self.__dict__.update(settings_dict)
        atexit.register(self.cleanup)
        self._excel_dir.mkdir(exist_ok=True, parents=True)
        self._extracted_files_dir.mkdir(exist_ok=True, parents=True)

    def __enter__(self):
        return self

    @staticmethod
    def find_docs_path():
        # POL_TAX_HOME overrides the data folder (used by tests)
        override = os.environ.get("POL_TAX_HOME")
        if override:
            path = Path(override)
            path.mkdir(parents=True, exist_ok=True)
            return path

        candidates = [
            QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation),
            QStandardPaths.writableLocation(QStandardPaths.GenericDataLocation),
            str(Path.home()),
        ]
        for candidate in candidates:
            if not candidate:
                continue
            path = Path(candidate) / "pol-tax"
            try:
                path.mkdir(parents=True, exist_ok=True)
                test_file = path / ".write_test"
                test_file.touch(exist_ok=True)
                test_file.unlink()
                return path
            except OSError:
                logging.warning(f"Folder {path} is not writable")
        logging.error("No path is accessible!")
        return Path.home() / "pol-tax"

    def get_excel_dir(self):
        return str(self._excel_dir.absolute())

    def get_extracted_dir(self):
        return self._extracted_files_dir.absolute()

    def get_keywords(self):
        return self._additional_keywords

    def get_docs_path(self):
        return self._docs_path

    def set_keywords(self, d):
        self._additional_keywords = d

    def set_excel_dir(self, path):
        self._excel_dir = Path(path)
        self._excel_dir.mkdir(exist_ok=True, parents=True)

    def cleanup(self):
        with self._settings_path.open(mode='wb') as output_file:
            pickle.dump(self.__dict__, output_file)

    def __exit__(self, exc_type, exc_value, traceback):
        self.cleanup()
