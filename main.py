from PySide6.QtWidgets import QApplication, QWidget, QStyleFactory
from PySide6.QtGui import QColor, QPalette
from PySide6.QtCore import QObject, Signal

from Widgets.MainWindow import MainGUI

from Settings import Settings

import sys
import subprocess

import logging

settings = Settings()


loging_level = logging.DEBUG
# Set up the logger
logger = logging.getLogger()  # Gets the root logger
logger.setLevel(loging_level)  # Set the minimum logging level to log

# Create file handler which logs warnings and above
file_handler = logging.FileHandler(f'{settings.get_docs_path()}/app.log')
file_handler.setLevel(loging_level)

# Create console handler with a higher log level
console_handler = logging.StreamHandler()
console_handler.setLevel(loging_level)

# Create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Add the handlers to the logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)


class ExceptHook(QObject):
    error = Signal(object)

    def handle_exception(self, exctype, value, traceback):
        self.error.emit((exctype, value, traceback))


def handle_exception(error_info):
    exctype, value, traceback = error_info
    logger.debug("handle_exception is being triggered!")
    logger.error("Uncaught Exception:",
                  exc_info=(exctype, value, traceback))


if __name__ == "__main__":
    excepthook = ExceptHook()
    sys.excepthook = excepthook.handle_exception

    # Log a test message
    logger.debug("App Started - Begging of Logging!")
    logger.debug(f"Current doc save path: {settings.get_docs_path()}")

    app = QApplication(sys.argv)

    excepthook.error.connect(handle_exception)
    logging.captureWarnings(True)
    app.setStyle(QStyleFactory.create('Material'))

    # try:
    #     settings = Settings()
    # except PermissionError as e:
    #     logger.error("Permission Error:", exc_info=(type(e), e, e.__traceback__))
    #     subprocess.run([sys.executable] + sys.argv, shell=True, elevation=True)
    #     sys.exit(0)

    window = MainGUI(app, settings)
    window.show()

    sys.exit(app.exec())

