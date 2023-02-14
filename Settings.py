import pickle
from pathlib import Path
import atexit


class Settings:
    _excel_dir = Path("./Excel")
    _extracted_files_dir = Path("./Extracted")
    # TODO: Paradox of opening setting file that can change it's location after closing
    _settings_path = Path("./settings.pickle")
    _additional_keywords = {
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

        if self._settings_path.exists():
            with self._settings_path.open(mode='rb') as settings_file:
                settings_dict = pickle.load(settings_file)
                self.__dict__.update(settings_dict)
        atexit.register(self.cleanup)
        self._excel_dir.mkdir(exist_ok=True, parents=True)
        self._extracted_files_dir.mkdir(exist_ok=True, parents=True)

    def __enter__(self):
        return self

    def get_excel_dir(self):
        return str(self._excel_dir.absolute())

    def get_extracted_dir(self):
        return self._extracted_files_dir.absolute()

    def get_keywords(self):
        return self._additional_keywords

    def set_excel_dir(self, path):
        self._excel_dir = Path(path)

    def cleanup(self):
        with self._settings_path.open(mode='wb') as output_file:
            pickle.dump(self.__dict__, output_file)

    def __exit__(self, exc_type, exc_value, traceback):
        print("Dumping Settings")
        with self._settings_path.open(mode='wb') as output_file:
            pickle.dump(self.__dict__, output_file)
