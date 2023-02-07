from PySide6.QtWidgets import QMainWindow, QFileDialog

from pathlib import Path

from XML.FinancialXMLParser import FinancialXMLParser


class MainWindow(QMainWindow):
    def __init__(self, app, app_settings):
        super().__init__()
        self.app = app
        self.app_settings = app_settings
        self.setWindowTitle("Pol-Tax")

        self.xml_filepath = Path("./").resolve()
        self.setWindowTitle("Pol-Tax")

        # Menu bar
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("&Plik")

        xml_file_action = file_menu.addAction("Otwórz XML")
        xml_file_action.triggered.connect(self.xml_file_dialog_open)

        quit_action = file_menu.addAction("Zamknij")
        quit_action.triggered.connect(self.quit_app)

    def quit_app(self):
        self.app.quit()

    def xml_file_dialog_open(self):
        filepath = QFileDialog.getOpenFileName(self, 'Wybierz plik XML', self.get_last_file_path_dir(),
                                                   "Image files (*.xml)")
        self.xml_filepath = Path(filepath[0])
        print(filepath)
        f_xml = FinancialXMLParser(str(self.xml_filepath), self.app_settings)
        f_xml.list_all_root_tags()
        f_xml.generate_assets_excel(f'{self.app_settings.get_excel_dir()}/Aktywa.xlsx')
        f_xml.generate_passives_excel(f'{self.app_settings.get_excel_dir()}/Pasywa.xlsx')

    def get_last_file_path_dir(self):
        print(self.xml_filepath.resolve().parents[0])
        return str(self.xml_filepath.resolve().parents[0])