import pickle
from pathlib import Path


class Settings:
    _excel_dir = Path("./Excel")
    # TODO: Paradox of opening setting file that can change it's location after closing
    _settings_path = Path("./settings.pickle")
    _additional_keywords = {
        "DataOd": "Okres od",
        "DataDo": "Okres do",
        "NazwaFirmy": "Nazwa Firmy",
        "Siedziba": "Siedziba",
        "P_1B": "",
        "P_1D": "NIP",
        "P_1E": "KRS"
    }

    def __init__(self):
        if self._settings_path.exists():
            with self._settings_path.open(mode='rb') as settings_file:
                settings_dict = pickle.load(settings_file)
                self.__dict__.update(settings_dict)

        self._excel_dir.mkdir(exist_ok=True, parents=True)

    def __enter__(self):
        return self

    def get_excel_dir(self):
        return str(self._excel_dir.absolute())

    def get_keywords(self):
        return self._additional_keywords

    def __exit__(self, exc_type, exc_value, traceback):
        with self._settings_path.open(mode='wb') as output_file:
            pickle.dump(self.__dict__, output_file)
