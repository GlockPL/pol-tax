from PySide6.QtWidgets import QApplication, QWidget
from Widgets.MainWidget import MainWidget
from Widgets.MainWindow import MainWindow

import sys

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # window = MainWidget()
    window = MainWindow(app)
    window.show()

    app.exec()
