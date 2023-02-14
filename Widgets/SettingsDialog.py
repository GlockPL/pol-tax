from pathlib import Path
from PySide6 import QtWidgets, QtGui
from Settings import Settings


class SettingsDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ustawienia")
        self.setWindowIcon(QtGui.QIcon("./resources/baseline_settings_black_48dp.png"))
        self.settings = Settings()

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
        # layout.addWidget(server_url_label)
        # layout.addWidget(self.server_url_lineedit)
        # layout.addWidget(api_key_label)
        # layout.addWidget(self.api_key_lineedit)
        # layout.addWidget(cache_enabled_label)
        # layout.addWidget(self.cache_enabled_checkbox)
        layout.addWidget(output_path_label)
        output_path_layout = QtWidgets.QHBoxLayout()
        output_path_layout.addWidget(self.output_path_lineedit)
        output_path_layout.addWidget(output_path_button)
        layout.addLayout(output_path_layout)
        self.setLayout(layout)

        # Add OK and Cancel buttons
        buttons_layout = QtWidgets.QHBoxLayout()
        ok_button = QtWidgets.QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        cancel_button = QtWidgets.QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        buttons_layout.addWidget(ok_button)
        buttons_layout.addWidget(cancel_button)
        layout.addLayout(buttons_layout)

    def browse_output_path(self):
        options = QtWidgets.QFileDialog.Options()
        output_dir = QtWidgets.QFileDialog.getExistingDirectory(self, "Ścieżka do folderu", options=options)
        if output_dir:
            self.output_path_lineedit.setText(output_dir)
            self.settings.set_excel_dir(output_dir)

    def get_settings(self):
        return Path(self.output_path_lineedit.text())

