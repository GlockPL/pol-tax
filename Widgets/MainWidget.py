from PySide6.QtWidgets import QWidget, QVBoxLayout, QPushButton, QFileDialog
from pathlib import Path

from XML.FinancialXMLParser import FinancialXMLParser


class MainWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.xml_filepath = Path("./").resolve()
        self.setWindowTitle("Pol-Tax")

        self.file_dialog_button = QPushButton("Znajdź XML")
        self.file_dialog_button.clicked.connect(self.button_file_dialog_clicked)

        self.vbox_layout = QVBoxLayout()
        self.vbox_layout.addWidget(self.file_dialog_button)

        self.setLayout(self.vbox_layout)

    def button_file_dialog_clicked(self):
        filepath = QFileDialog.getOpenFileName(self, 'Wybierz plik XML', self.get_last_file_path_dir(),
                                                   "Image files (*.xml)")
        self.xml_filepath = Path(filepath[0])
        print(filepath)
        f_xml = FinancialXMLParser(str(self.xml_filepath))
        f_xml.lit_all_root_tags()
        f_xml.generate_assets_excel('Aktywa.xlsx')
        f_xml.generate_passives_excel('Pasywa.xlsx')

    def get_last_file_path_dir(self):
        print(self.xml_filepath.resolve().parents[0])
        return str(self.xml_filepath.resolve().parents[0])