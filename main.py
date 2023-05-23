from PySide6.QtWidgets import QApplication, QWidget, QStyleFactory
from PySide6.QtGui import QColor, QPalette
from PySide6.QtCore import QObject, Signal

from Widgets.MainWindow import MainGUI

from Settings import Settings

import sys
import subprocess

import logging

# Set up the logger
logging.basicConfig(filename='app.log', level=logging.DEBUG,
                    format='%(asctime)s %(levelname)s %(name)s %(message)s')


class ExceptHook(QObject):
    error = Signal(object)

    def handle_exception(self, exctype, value, traceback):
        self.error.emit((exctype, value, traceback))


def handle_exception(error_info):
    exctype, value, traceback = error_info
    print("handle_exception is being triggered!")
    logging.error("Uncaught Exception:",
                  exc_info=(exctype, value, traceback))


if __name__ == "__main__":
    excepthook = ExceptHook()
    sys.excepthook = excepthook.handle_exception

    # Log a test message
    logging.warning("App Started - Begging of Logging!")

    app = QApplication(sys.argv)

    excepthook.error.connect(handle_exception)
    logging.captureWarnings(True)
    app.setStyle(QStyleFactory.create('Material'))

    # try:
    settings = Settings()
    # except PermissionError as e:
    #     logging.error("Permission Error:", exc_info=(type(e), e, e.__traceback__))
    #     subprocess.run([sys.executable] + sys.argv, shell=True, elevation=True)
    #     sys.exit(0)

    window = MainGUI(app, settings)
    window.show()

    sys.exit(app.exec())
