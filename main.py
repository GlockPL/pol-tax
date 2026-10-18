from PySide6.QtWidgets import QApplication, QMessageBox, QStyleFactory
from PySide6.QtCore import QObject, Signal

from Widgets.MainWindow import MainGUI

from Settings import Settings

import sys

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
    QMessageBox.critical(None, "Błąd",
                         f"Wystąpił nieoczekiwany błąd:\n\n{value}\n\n"
                         f"Szczegóły zapisano w {settings.get_docs_path() / 'app.log'}")


if __name__ == "__main__":
    excepthook = ExceptHook()
    sys.excepthook = excepthook.handle_exception

    # Log a test message
    logger.debug("App Started - Begging of Logging!")
    logger.debug(f"Current doc save path: {settings.get_docs_path()}")

    app = QApplication(sys.argv)

    excepthook.error.connect(handle_exception)
    logging.captureWarnings(True)
    app.setStyle(QStyleFactory.create('Fusion'))

    window = MainGUI(app, settings)
    window.show()

    # "Open with" from Explorer passes the file as the first argument
    if len(sys.argv) > 1:
        window.open_file(sys.argv[1])

    sys.exit(app.exec())

