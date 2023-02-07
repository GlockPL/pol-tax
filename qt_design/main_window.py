# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'main_window.ui'
##
## Created by: Qt User Interface Compiler version 6.4.2
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QAction, QBrush, QColor, QConicalGradient,
    QCursor, QFont, QFontDatabase, QGradient,
    QIcon, QImage, QKeySequence, QLinearGradient,
    QPainter, QPalette, QPixmap, QRadialGradient,
    QTransform)
from PySide6.QtWidgets import (QApplication, QHeaderView, QLabel, QMainWindow,
    QMenu, QMenuBar, QSizePolicy, QStatusBar,
    QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(878, 631)
        icon = QIcon(QIcon.fromTheme(u"accessories-dictionary"))
        MainWindow.setWindowIcon(icon)
        self.actionOtw_rz_XML = QAction(MainWindow)
        self.actionOtw_rz_XML.setObjectName(u"actionOtw_rz_XML")
        self.actionZamknij = QAction(MainWindow)
        self.actionZamknij.setObjectName(u"actionZamknij")
        self.actionO_programie = QAction(MainWindow)
        self.actionO_programie.setObjectName(u"actionO_programie")
        self.widget = QWidget(MainWindow)
        self.widget.setObjectName(u"widget")
        self.widget1 = QWidget(self.widget)
        self.widget1.setObjectName(u"widget1")
        self.widget1.setGeometry(QRect(20, 30, 491, 341))
        self.verticalLayout = QVBoxLayout(self.widget1)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.label = QLabel(self.widget1)
        self.label.setObjectName(u"label")

        self.verticalLayout.addWidget(self.label)

        self.treeWidget = QTreeWidget(self.widget1)
        QTreeWidgetItem(self.treeWidget)
        QTreeWidgetItem(self.treeWidget)
        QTreeWidgetItem(self.treeWidget)
        QTreeWidgetItem(self.treeWidget)
        self.treeWidget.setObjectName(u"treeWidget")

        self.verticalLayout.addWidget(self.treeWidget)

        MainWindow.setCentralWidget(self.widget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName(u"menubar")
        self.menubar.setGeometry(QRect(0, 0, 878, 22))
        self.menuPlik = QMenu(self.menubar)
        self.menuPlik.setObjectName(u"menuPlik")
        self.menuPomoc = QMenu(self.menubar)
        self.menuPomoc.setObjectName(u"menuPomoc")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.menubar.addAction(self.menuPlik.menuAction())
        self.menubar.addAction(self.menuPomoc.menuAction())
        self.menuPlik.addAction(self.actionOtw_rz_XML)
        self.menuPlik.addAction(self.actionZamknij)
        self.menuPomoc.addSeparator()
        self.menuPomoc.addAction(self.actionO_programie)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"Pol-Tax", None))
        self.actionOtw_rz_XML.setText(QCoreApplication.translate("MainWindow", u"Otw\u00f3rz XML", None))
        self.actionZamknij.setText(QCoreApplication.translate("MainWindow", u"Zamknij", None))
        self.actionO_programie.setText(QCoreApplication.translate("MainWindow", u"O programie", None))
        self.label.setText(QCoreApplication.translate("MainWindow", u"TextLabel", None))
        ___qtreewidgetitem = self.treeWidget.headerItem()
        ___qtreewidgetitem.setText(4, QCoreApplication.translate("MainWindow", u"Weryfikuj\u0105cy", None));
        ___qtreewidgetitem.setText(3, QCoreApplication.translate("MainWindow", u"Sprawdzaj\u0105cy", None));
        ___qtreewidgetitem.setText(2, QCoreApplication.translate("MainWindow", u"Wykonuj\u0105cy", None));
        ___qtreewidgetitem.setText(1, QCoreApplication.translate("MainWindow", u"W\u0142a\u015bciciel", None));
        ___qtreewidgetitem.setText(0, QCoreApplication.translate("MainWindow", u"Dokument", None));

        __sortingEnabled = self.treeWidget.isSortingEnabled()
        self.treeWidget.setSortingEnabled(False)
        ___qtreewidgetitem1 = self.treeWidget.topLevelItem(0)
        ___qtreewidgetitem1.setText(0, QCoreApplication.translate("MainWindow", u"A. Dokumentacja Sta\u0142a", None));
        ___qtreewidgetitem2 = self.treeWidget.topLevelItem(1)
        ___qtreewidgetitem2.setText(0, QCoreApplication.translate("MainWindow", u"B. Karta identyfikacyjna", None));
        ___qtreewidgetitem3 = self.treeWidget.topLevelItem(2)
        ___qtreewidgetitem3.setText(0, QCoreApplication.translate("MainWindow", u"C. Planowanie", None));
        ___qtreewidgetitem4 = self.treeWidget.topLevelItem(3)
        ___qtreewidgetitem4.setText(0, QCoreApplication.translate("MainWindow", u"D. Kontrola Wewn\u0119trzna", None));
        self.treeWidget.setSortingEnabled(__sortingEnabled)

        self.menuPlik.setTitle(QCoreApplication.translate("MainWindow", u"Plik", None))
        self.menuPomoc.setTitle(QCoreApplication.translate("MainWindow", u"Pomoc", None))
    # retranslateUi

