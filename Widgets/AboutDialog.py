from PySide6 import QtWidgets, QtCore
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import QDialog, QDialogButtonBox, QHBoxLayout

from Other.Version import Version

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        version = Version()

        self.setWindowTitle("O programie")
        self.setWindowIcon(QIcon(QPixmap(":/icons/pol_tax_135_icon.png")))
        self.resize(390, 220)

        horizontalLayout = QHBoxLayout()

        image_label = QtWidgets.QLabel()
        pixmap = QPixmap(":/icons/pol_tax_540_icon.jpg")
        pixmap = pixmap.scaled(135, 135, QtCore.Qt.KeepAspectRatio)
        image_label.setPixmap(pixmap)
        image_label.setAlignment(QtCore.Qt.AlignCenter)

        horizontalLayout.addWidget(image_label)

        # Label for displaying information
        info_label = QtWidgets.QLabel(f"Pol-Tax \n Wersja {version} \nCopyright © 2023\n Stworzona przy pomocy PySide6 i Qt")
        info_label .setAlignment(QtCore.Qt.AlignCenter)

        horizontalLayout.addWidget(info_label)

        buttonBox = QDialogButtonBox()
        buttonBox.setObjectName(u"buttonBox")
        buttonBox.setOrientation(QtCore.Qt.Horizontal)
        buttonBox.setStandardButtons(QDialogButtonBox.Ok)
        buttonBox.setCenterButtons(True)
        buttonBox.accepted.connect(self.accept)

        # Button for closing the dialog
        # close_button = QtWidgets.QPushButton("Zamknij")
        # close_button.clicked.connect(self.close)

        # Layout for the dialog
        grid = QtWidgets.QGridLayout()
        grid.addLayout(horizontalLayout, 0, 0, 1, 1)
        # grid.addWidget(image_label, 0, 0)
        # grid.addWidget(info_label, 0, 1)
        grid.addWidget(buttonBox, 1, 0, 1, 1)
        # grid.addWidget(buttonBox, 1, 1)
        self.setLayout(grid)

