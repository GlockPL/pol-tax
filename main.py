from PySide6.QtWidgets import QApplication, QWidget
from PySide6.QtGui import QColor, QPalette
from Widgets.MainWindow import MainGUI

from Settings import Settings

import sys

if __name__ == "__main__":
    # Define your colors
    background_color = QColor("#F2F2F2")
    text_color = QColor("#000000")
    accent_color = QColor("#0078D7")
    border_color = QColor("#C2C2C2")

    # Create a palette and set the colors
    palette = QPalette()
    palette.setColor(QPalette.Window, background_color)
    palette.setColor(QPalette.WindowText, text_color)
    palette.setColor(QPalette.Button, background_color)
    palette.setColor(QPalette.ButtonText, text_color)
    palette.setColor(QPalette.ToolTipBase, background_color)
    palette.setColor(QPalette.ToolTipText, text_color)
    palette.setColor(QPalette.Link, accent_color)
    palette.setColor(QPalette.LinkVisited, accent_color)
    palette.setColor(QPalette.Disabled, QPalette.Text, QColor("#999999"))
    palette.setColor(QPalette.Disabled, QPalette.WindowText, QColor("#999999"))
    palette.setColor(QPalette.Disabled, QPalette.ButtonText, QColor("#999999"))
    palette.setColor(QPalette.Disabled, QPalette.Highlight, QColor("#DDDDDD"))
    palette.setColor(QPalette.Disabled, QPalette.HighlightedText, QColor("#999999"))
    palette.setColor(QPalette.Mid, border_color)
    palette.setColor(QPalette.Dark, QColor("#9B9B9B"))

    # Set the base color to a light gray color
    palette.setColor(QPalette.Base, QColor('#fbfbfb'))
    # Set the text color to black
    palette.setColor(QPalette.Text, QColor('#000000'))
    # Set the highlight color to a light blue color
    palette.setColor(QPalette.Highlight, QColor('#A3C7FF'))
    # Set the highlighted text color to white
    palette.setColor(QPalette.HighlightedText, QColor('#FFFFFF'))

    app = QApplication(sys.argv)
    # app.setPalette(palette)
    app.setStyle('Fusion')
    settings = Settings()

    window = MainGUI(app, settings)
    window.show()

    app.exec()
