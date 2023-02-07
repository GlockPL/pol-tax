from PySide6.QtWidgets import QApplication, QWidget
from Widgets.MainWidget import MainWidget
from Widgets.MainWindow import MainGUI

from Settings import Settings

import sys

if __name__ == "__main__":
    app = QApplication(sys.argv)
    settings = Settings()

    # window = MainWidget()
    window = MainGUI(app, settings)
    window.show()

    app.exec()
