import logging

from PySide6.QtCore import QStandardPaths, QDir

import pickle
import os
from pathlib import Path
import atexit


class Settings:
    _excel_dir = Path("./Excel")
    _extracted_files_dir = Path("./Extracted")
    # TODO: Paradox of opening setting file that can change it's location after closing
    _settings_path = Path("./settings.pickle")
    _docs_path = Path(QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)) / "pol-tax"
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
        self.test_permissible_paths()

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

    def test_permissible_paths(self):
        directories = [
            Path.home() / "Documents",
            Path.home() / "Desktop",
        ]

        # Add all available drives to the list
        drives = os.popen('wmic logicaldisk get name').read().split()[1:]
        directories.extend(Path(drive) for drive in drives)

        for directory in directories:
            if directory.exists():
                try:
                    test_file = directory / "test_file.tmp"
                    test_file.touch(exist_ok=True)
                    test_file.unlink()  # remove the test file
                    self._docs_path = directory / "pol-tax"
                    return
                except PermissionError:
                    pass
        logging.error("No path is accessible!")

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

    def cleanup(self):
        with self._settings_path.open(mode='wb') as output_file:
            pickle.dump(self.__dict__, output_file)

    def __exit__(self, exc_type, exc_value, traceback):
        print("Dumping Settings")
        with self._settings_path.open(mode='wb') as output_file:
            pickle.dump(self.__dict__, output_file)
