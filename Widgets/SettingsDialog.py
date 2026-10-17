from pathlib import Path
from PySide6 import QtWidgets
from PySide6.QtWidgets import QTableView, QHeaderView, QLabel
from PySide6.QtGui import QStandardItemModel, QStandardItem, QIcon, QPixmap
from Settings import Settings
import Widgets.icons  # noqa: F401  - registers the icons: search path


class SettingsDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ustawienia")
        self.setWindowIcon(QIcon(QPixmap("icons:baseline_settings_black_48dp.png")))
        self.settings = Settings()
        self.setMinimumWidth(390)
        self.setMinimumHeight(455)

        # Create widgets for each field of the dataclass
        # server_url_label = QtWidgets.QLabel("Server URL:")
        # self.server_url_lineedit = QtWidgets.QLineEdit(settings.server_url)
        # api_key_label = QtWidgets.QLabel("API Key:")
        # self.api_key_lineedit = QtWidgets.QLineEdit(settings.api_key)
        # cache_enabled_label = QtWidgets.QLabel("Cache Enabled:")
        # self.cache_enabled_checkbox = QtWidgets.QCheckBox()
        # self.cache_enabled_checkbox.setChecked(settings.cache_enabled)
        output_path_label = QtWidgets.QLabel("Ścieżka zapisu plików Excel:")
        self.output_path_lineedit = QtWidgets.QLineEdit(str(self.settings.get_excel_dir()))
        output_path_button = QtWidgets.QPushButton("Przeglądaj...")
        output_path_button.clicked.connect(self.browse_output_path)

        # Add the widgets to a layout
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(output_path_label)
        output_path_layout = QtWidgets.QHBoxLayout()
        output_path_layout.addWidget(self.output_path_lineedit)
        output_path_layout.addWidget(output_path_button)
        layout.addLayout(output_path_layout)
        # Settings dict as table
        self.table_view = QTableView()

        # Create the model for the table view
        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(['Klucz', 'Wartość'])

        # Add the data from the dictionary to the model
        for key, value in self.settings.get_keywords().items():
            key_item = QStandardItem(key)
            key_item.setEditable(False)
            value_item = QStandardItem(str(value))
            self.model.appendRow([key_item, value_item])

        # Set the model for the table view
        self.table_view.setModel(self.model)
        horizontalHead = self.table_view.horizontalHeader()
        horizontalHead.setSectionResizeMode(QHeaderView.ResizeToContents)

        internal_layout = QtWidgets.QVBoxLayout()
        tabel_label = QLabel("Dodatkowe Klucze:")
        internal_layout.addWidget(tabel_label)
        internal_layout.addWidget(self.table_view)

        layout.addLayout(internal_layout)
        self.setLayout(layout)

        # Add OK and Cancel buttons
        buttons_layout = QtWidgets.QHBoxLayout()
        ok_button = QtWidgets.QPushButton("Zapisz")
        ok_button.clicked.connect(self.save_settings)
        cancel_button = QtWidgets.QPushButton("Anuluj")
        cancel_button.clicked.connect(self.reject)
        buttons_layout.addWidget(ok_button)
        buttons_layout.addWidget(cancel_button)
        layout.addLayout(buttons_layout)

    def browse_output_path(self):
        options = QtWidgets.QFileDialog.Options()
        output_dir = QtWidgets.QFileDialog.getExistingDirectory(self, "Ścieżka do folderu", options=options)
        if output_dir:
            self.output_path_lineedit.setText(output_dir)

    def save_settings(self):
        self.settings.set_excel_dir(self.output_path_lineedit.text())
        additional_keywords_dict = {}
        for row in range(self.model.rowCount()):
            key = self.model.item(row, 0).text()
            value = self.model.item(row, 1).text()
            additional_keywords_dict[key] = value

        self.settings.set_keywords(additional_keywords_dict)
        self.accept()

    def get_settings(self):
        return Path(self.output_path_lineedit.text())

