import logging
import os
from pathlib import Path

from PySide6.QtCore import QSize, QStandardPaths, Qt, QThreadPool, QUrl
from PySide6.QtGui import QAction, QColor, QDesktopServices, QIcon, QKeySequence, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox, QDialog, QFileDialog, QFrame,
                               QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
                               QMenu,
                               QMessageBox, QProgressBar, QPushButton, QSplitter, QStackedWidget,
                               QTableWidget, QTableWidgetItem, QToolBar, QTreeWidget, QTreeWidgetItem, QVBoxLayout,
                               QWidget)

import Widgets.icons  # noqa: F401  - registers the icons: search path
from Model.PandasModel import PandasModel, format_amount
from Widgets.AboutDialog import AboutDialog
from Widgets.BackgroundTask import BackgroundTask
from Widgets.ColorDelegate import ColorDelegate
from Widgets.SettingsDialog import SettingsDialog
from Widgets.StatementTable import StatementTable
from XML.export import (SECTIONS, ExportResult, SummationResult, default_workbook_name, export_statement,
                        export_summation, export_workbook)
from XML.presentation import (CHANGE_COLUMN, CHANGE_PERCENT_COLUMN, YEAR_SECTIONS, short_year_headers,
                              with_year_change)
from XML.rules import SECTION_TITLES, YEAR_NAMES

RDF_URL = "https://ekrs.ms.gov.pl/rdf/pd/search_df"
FAILED_TEXT_COLOR = QColor("#b02a37")
OK_TEXT_COLOR = QColor("#2e7d32")


def plural(count, one, few, many):
    """
    Polish plural: 1 reguła, 2 reguły, 5 reguł
    """
    if count == 1:
        return f"{count} {one}"
    if count % 10 in (2, 3, 4) and count % 100 not in (12, 13, 14):
        return f"{count} {few}"
    return f"{count} {many}"


def xml_paths_from_mime(mime_data):
    """
    Local .xml files from a drag and drop
    """
    if not mime_data.hasUrls():
        return []
    paths = [Path(url.toLocalFile()) for url in mime_data.urls() if url.isLocalFile()]
    return [path for path in paths if path.suffix.lower() == ".xml"]


class MainGUI(QMainWindow):
    def __init__(self, app, app_settings):
        super().__init__()
        self.app = app
        self.app_settings = app_settings

        self.xml_filepath = None
        self.export: ExportResult | None = None
        self.summation: SummationResult | None = None
        self.task = None
        self.current_section = None
        self.current_model = None
        self.section_items = {}
        # Folded positions per section, kept while switching between sections
        self.collapsed_by_section = {}
        self.last_export_dir = None

        self.excel_icon = QIcon(QPixmap("icons:excel_icon.jpg"))
        self.spreadsheet_icon = QIcon(QPixmap("icons:baseline_table_view_black_48dp.png"))
        self.pdf_icon = QIcon(QPixmap("icons:baseline_picture_as_pdf_black_48dp.png"))
        self.main_icon = QIcon(QPixmap("icons:cropped-favicon-1-270x270.png"))
        self.check_icon = self.load_svg_and_color("icons:summarize_black_48dp.svg", '#5f6368')
        self.folder_icon = self.load_svg_and_color("icons:folder_open_black_48dp.svg", '#5f6368')

        self.setWindowIcon(self.main_icon)
        self.resize(1280, 820)
        self.setAcceptDrops(True)

        self.create_actions()
        self.create_menus()
        self.create_toolbar()
        self.create_status_bar()

        self.stack = QStackedWidget()
        self.welcome_page = self.create_welcome_page()
        self.document_page = self.create_document_page()
        self.stack.addWidget(self.welcome_page)
        self.stack.addWidget(self.document_page)
        self.setCentralWidget(self.stack)

        self.update_actions()
        self.update_title()

    # ------------------------------------------------------------------ setup

    def create_actions(self):
        self.actionOpenXML = QAction(self.style().standardIcon(self.style().StandardPixmap.SP_DialogOpenButton),
                                     "Otwórz XML…", self)
        self.actionOpenXML.setShortcut(QKeySequence.Open)
        self.actionOpenXML.triggered.connect(self.xml_file_dialog_open)

        self.actionCheck = QAction(self.check_icon, "Sprawdź sumy", self)
        self.actionCheck.setShortcut("F5")
        self.actionCheck.setToolTip("Sprawdź zgodność sum i powiązań między częściami sprawozdania (F5)")
        self.actionCheck.triggered.connect(self.xml_summation_check)

        self.actionOpenExcel = QAction(self.excel_icon, "Otwórz w Excelu", self)
        self.actionOpenExcel.setShortcut("Ctrl+E")
        self.actionOpenExcel.setToolTip("Otwórz wyświetlaną część sprawozdania w Excelu (Ctrl+E)")
        self.actionOpenExcel.triggered.connect(self.open_current_excel)

        self.actionExport = QAction(self.style().standardIcon(self.style().StandardPixmap.SP_DialogSaveButton),
                                    "Eksportuj do Excela…", self)
        self.actionExport.setShortcut(QKeySequence.Save)
        self.actionExport.setToolTip("Zapisz całe sprawozdanie i wynik kontroli w jednym pliku Excel (Ctrl+S)")
        self.actionExport.triggered.connect(self.export_to_excel)

        self.actionFolder = QAction(self.folder_icon, "Folder z plikami", self)
        self.actionFolder.setToolTip("Otwórz folder z wygenerowanymi plikami Excel")
        self.actionFolder.triggered.connect(self.open_excel_folder)

        self.actionSettings = QAction("Ustawienia…", self)
        self.actionSettings.setShortcut(QKeySequence.Preferences)
        self.actionSettings.triggered.connect(self.show_settings)

        self.actionWhereXML = QAction("Skąd wziąć plik XML?", self)
        self.actionWhereXML.triggered.connect(lambda: QDesktopServices.openUrl(QUrl(RDF_URL)))

        self.actionAbout = QAction("O programie", self)
        self.actionAbout.triggered.connect(self.show_about)

        self.actionClose = QAction("Zamknij", self)
        self.actionClose.setShortcut(QKeySequence.Quit)
        self.actionClose.triggered.connect(self.quit_app)

    def create_menus(self):
        file_menu = self.menuBar().addMenu("&Plik")
        file_menu.addAction(self.actionOpenXML)
        self.recent_menu = file_menu.addMenu("Ostatnio otwierane")
        self.recent_menu.aboutToShow.connect(self.fill_recent_menu)
        file_menu.addSeparator()
        file_menu.addAction(self.actionExport)
        file_menu.addAction(self.actionOpenExcel)
        file_menu.addAction(self.actionFolder)
        file_menu.addSeparator()
        file_menu.addAction(self.actionClose)

        tools_menu = self.menuBar().addMenu("&Narzędzia")
        tools_menu.addAction(self.actionCheck)
        tools_menu.addSeparator()
        tools_menu.addAction(self.actionSettings)

        help_menu = self.menuBar().addMenu("Pomo&c")
        help_menu.addAction(self.actionWhereXML)
        help_menu.addSeparator()
        help_menu.addAction(self.actionAbout)

    def create_toolbar(self):
        toolbar = QToolBar("Pasek narzędzi", self)
        toolbar.setObjectName("mainToolbar")
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(20, 20))
        toolbar.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        toolbar.addAction(self.actionOpenXML)
        toolbar.addSeparator()
        toolbar.addAction(self.actionCheck)
        toolbar.addAction(self.actionExport)
        toolbar.addAction(self.actionOpenExcel)
        toolbar.addAction(self.actionFolder)
        self.addToolBar(Qt.TopToolBarArea, toolbar)

    def create_status_bar(self):
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setMaximumWidth(160)
        self.progress.setVisible(False)
        self.statusBar().addPermanentWidget(self.progress)

    def create_welcome_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.addStretch(2)

        logo = QLabel()
        logo.setPixmap(QPixmap("icons:cropped-favicon-1-270x270.png").scaled(96, 96, Qt.KeepAspectRatio,
                                                                          Qt.SmoothTransformation))
        logo.setAlignment(Qt.AlignCenter)
        layout.addWidget(logo)

        title = QLabel("Pol-Tax")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 26px; font-weight: 600;")
        layout.addWidget(title)

        hint = QLabel("Przeciągnij tutaj plik XML ze sprawozdaniem finansowym")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet("font-size: 15px;")
        layout.addWidget(hint)

        open_button = QPushButton("Otwórz plik XML…")
        open_button.setMinimumHeight(36)
        open_button.setMinimumWidth(220)
        open_button.clicked.connect(self.xml_file_dialog_open)
        button_row = QHBoxLayout()
        button_row.addStretch()
        button_row.addWidget(open_button)
        button_row.addStretch()
        layout.addSpacing(8)
        layout.addLayout(button_row)

        self.welcome_recent_label = QLabel("Ostatnio otwierane:")
        self.welcome_recent_label.setAlignment(Qt.AlignCenter)
        self.welcome_recent = QListWidget()
        self.welcome_recent.setMaximumWidth(620)
        self.welcome_recent.setMaximumHeight(150)
        self.welcome_recent.setFrameShape(QFrame.NoFrame)
        self.welcome_recent.itemClicked.connect(lambda item: self.open_file(item.data(Qt.UserRole)))
        recent_row = QHBoxLayout()
        recent_row.addStretch()
        recent_row.addWidget(self.welcome_recent)
        recent_row.addStretch()
        layout.addSpacing(16)
        layout.addWidget(self.welcome_recent_label)
        layout.addLayout(recent_row)
        self.fill_welcome_recent()

        where = QLabel(f'Skąd wziąć plik? Sprawozdania wszystkich spółek z KRS można pobrać w '
                       f'<a href="{RDF_URL}">Repozytorium Dokumentów Finansowych</a> (wyszukaj po numerze KRS, '
                       f'pobierz „Roczne sprawozdanie finansowe”, a jeżeli jest spakowane - rozpakuj plik .xml).')
        where.setWordWrap(True)
        where.setOpenExternalLinks(True)
        where.setAlignment(Qt.AlignCenter)
        where.setMaximumWidth(620)
        where_row = QHBoxLayout()
        where_row.addStretch()
        where_row.addWidget(where)
        where_row.addStretch()
        layout.addSpacing(16)
        layout.addLayout(where_row)

        layout.addStretch(3)
        return page

    def create_document_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(8, 8, 8, 4)

        self.company_label = QLabel()
        self.company_label.setStyleSheet("font-size: 18px; font-weight: 600;")
        self.company_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.details_label = QLabel()
        self.details_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(self.company_label)
        layout.addWidget(self.details_label)

        # Left: parts of the statement and attachments
        self.nav = QTreeWidget()
        self.nav.setHeaderHidden(True)
        self.nav.setMinimumWidth(220)
        self.nav.setContextMenuPolicy(Qt.CustomContextMenu)
        self.nav.customContextMenuRequested.connect(self.show_nav_menu)
        self.nav.currentItemChanged.connect(self.nav_item_changed)
        self.nav.itemDoubleClicked.connect(self.nav_item_activated)

        # Right: the table and below it the check results
        self.section_title = QLabel()
        self.section_title.setStyleSheet("font-size: 15px; font-weight: 600;")
        self.search = QLineEdit()
        self.search.setPlaceholderText("Szukaj pozycji (Ctrl+F)…")
        self.search.setClearButtonEnabled(True)
        self.search.setMaximumWidth(280)
        self.search.textChanged.connect(self.apply_row_visibility)
        search_action = QAction(self)
        search_action.setShortcut(QKeySequence.Find)
        search_action.triggered.connect(lambda: (self.search.setFocus(), self.search.selectAll()))
        self.addAction(search_action)
        self.expand_button = QPushButton("Rozwiń wszystko")
        self.expand_button.clicked.connect(self.expand_all)
        self.collapse_button = QPushButton("Zwiń")
        self.collapse_button.setToolTip("Pokaż tylko główne pozycje")
        self.collapse_button.clicked.connect(self.collapse_all)
        self.show_change = QCheckBox("Zmiana r/r")
        self.show_change.setToolTip("Pokaż zmianę względem poprzedniego roku")
        self.show_change.setChecked(True)
        self.show_change.toggled.connect(lambda: self.show_section(self.current_section))
        table_toolbar = QHBoxLayout()
        table_toolbar.addWidget(self.section_title)
        table_toolbar.addStretch()
        table_toolbar.addWidget(self.show_change)
        table_toolbar.addWidget(self.expand_button)
        table_toolbar.addWidget(self.collapse_button)
        table_toolbar.addWidget(self.search)

        self.tableView = StatementTable()
        self.tableView.verticalHeader().setDefaultSectionSize(26)
        self.tableView.toggle_requested.connect(self.toggle_row)
        table_box = QWidget()
        table_layout = QVBoxLayout(table_box)
        table_layout.setContentsMargins(0, 0, 0, 0)
        table_layout.addLayout(table_toolbar)
        table_layout.addWidget(self.tableView)

        self.check_panel = self.create_check_panel()
        self.check_panel.setVisible(False)

        right = QSplitter(Qt.Vertical)
        right.addWidget(table_box)
        right.addWidget(self.check_panel)
        right.setStretchFactor(0, 3)
        right.setStretchFactor(1, 1)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.nav)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([260, 1000])
        layout.addWidget(splitter, 1)
        return page

    def create_check_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 6, 0, 0)

        header = QHBoxLayout()
        self.check_summary = QLabel()
        self.only_failures = QCheckBox("Pokaż tylko niezgodności")
        self.only_failures.setChecked(True)
        self.only_failures.toggled.connect(self.fill_check_table)
        open_summation = QPushButton(self.excel_icon, "Raport w Excelu")
        open_summation.clicked.connect(self.open_summation_excel)
        header.addWidget(self.check_summary)
        header.addStretch()
        header.addWidget(self.only_failures)
        header.addWidget(open_summation)
        layout.addLayout(header)

        self.check_table = QTableWidget(0, 8)
        self.check_table.setHorizontalHeaderLabels(
            ["", "Część", "Pozycja", "Rok", "Wartość", "Oczekiwana", "Różnica", "Reguła (oczekiwana wartość)"])
        self.check_table.verticalHeader().setVisible(False)
        self.check_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.check_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.check_table.setWordWrap(False)
        header_view = self.check_table.horizontalHeader()
        for column in range(7):
            header_view.setSectionResizeMode(column, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(7, QHeaderView.Stretch)
        self.check_table.cellClicked.connect(self.check_row_clicked)
        layout.addWidget(self.check_table)
        return panel

    # ------------------------------------------------------------------ opening files

    def xml_file_dialog_open(self):
        start_dir = str(self.xml_filepath.parent) if self.xml_filepath else ""
        recent = self.app_settings.get_recent_files()
        if not start_dir and recent:
            start_dir = str(Path(recent[0]).parent)
        filepath, _ = QFileDialog.getOpenFileName(self, 'Wybierz plik XML', start_dir, "Pliki xml (*.xml)")
        if filepath:
            self.open_file(filepath)

    def open_file(self, filepath):
        path = Path(filepath)
        if path.suffix.lower() != '.xml':
            QMessageBox.warning(self, "Pol-Tax", f"To nie jest plik XML:\n{path}")
            return
        if not path.exists():
            QMessageBox.warning(self, "Pol-Tax", f"Plik nie istnieje:\n{path}")
            return
        logging.debug(f"Opening {path}")
        self.xml_filepath = path
        self.setup_excels()

    def setup_excels(self):
        if self.xml_filepath is None:
            return
        self.run_in_background(f"Wczytywanie {self.xml_filepath.name}…", self.show_export,
                               export_statement, self.xml_filepath, self.app_settings)

    def dragEnterEvent(self, event):
        if xml_paths_from_mime(event.mimeData()):
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = xml_paths_from_mime(event.mimeData())
        if paths:
            event.acceptProposedAction()
            self.open_file(paths[0])

    def fill_recent_menu(self):
        self.recent_menu.clear()
        recent = self.app_settings.get_recent_files()
        if not recent:
            action = self.recent_menu.addAction("(brak)")
            action.setEnabled(False)
            return
        for path in recent:
            action = self.recent_menu.addAction(Path(path).name)
            action.setStatusTip(path)
            action.triggered.connect(lambda checked=False, p=path: self.open_file(p))

    def fill_welcome_recent(self):
        self.welcome_recent.clear()
        recent = self.app_settings.get_recent_files()
        for path in recent:
            item = QListWidgetItem(f"{Path(path).name}   —   {Path(path).parent}")
            item.setData(Qt.UserRole, path)
            item.setToolTip(path)
            self.welcome_recent.addItem(item)
        self.welcome_recent.setVisible(bool(recent))
        self.welcome_recent_label.setVisible(bool(recent))

    # ------------------------------------------------------------------ background work

    def run_in_background(self, message, on_finished, fn, *args):
        if self.task is not None:
            self.statusBar().showMessage("Poczekaj na zakończenie przetwarzania", 2000)
            return
        self.task = BackgroundTask(fn, *args)
        self.set_busy(True, message)
        self.task.setAutoDelete(False)
        self.task.signals.finished.connect(lambda result: self.task_done(on_finished, result))
        self.task.signals.failed.connect(self.task_failed)
        QThreadPool.globalInstance().start(self.task)

    def set_busy(self, busy, message=""):
        self.progress.setVisible(busy)
        if busy:
            self.statusBar().showMessage(message)
            QApplication.setOverrideCursor(Qt.WaitCursor)
        else:
            self.statusBar().clearMessage()
            QApplication.restoreOverrideCursor()
        self.update_actions()

    def task_done(self, on_finished, result):
        self.task = None
        self.set_busy(False)
        on_finished(result)

    def task_failed(self, error):
        self.task = None
        self.set_busy(False)
        logging.error(f"Processing failed:\n{getattr(error, 'formatted_traceback', error)}")
        QMessageBox.critical(self, "Błąd", f"Nie udało się przetworzyć pliku:\n\n{error}")

    def update_actions(self):
        busy = self.task is not None
        loaded = self.export is not None
        self.actionOpenXML.setEnabled(not busy)
        self.actionSettings.setEnabled(not busy)
        self.actionCheck.setEnabled(loaded and not busy)
        self.actionOpenExcel.setEnabled(loaded and self.current_section in self.export.sections if loaded
                                        else False)
        self.actionExport.setEnabled(loaded and not busy)

    # ------------------------------------------------------------------ showing the document

    def show_export(self, result: ExportResult):
        self.export = result
        self.summation = None
        self.app_settings.add_recent_file(result.xml_path)
        self.fill_welcome_recent()

        info = result.parser.company_info()
        self.company_label.setText(info["name"] or result.xml_path.stem)
        details = []
        if info["nip"]:
            details.append(f"NIP {info['nip']}")
        if info["krs"]:
            details.append(f"KRS {info['krs']}")
        if info["city"]:
            details.append(info["city"])
        if info["date_from"] or info["date_to"]:
            details.append(f"okres {info['date_from']} – {info['date_to']}")
        details.append(result.xml_path.name)
        self.details_label.setText("   ·   ".join(details))

        self.fill_nav()
        self.check_panel.setVisible(False)
        self.stack.setCurrentWidget(self.document_page)
        self.update_title()
        first_key = "assets" if "assets" in result.sections else next(iter(result.sections), None)
        self.select_section(first_key)
        self.statusBar().showMessage(f"Wczytano {result.xml_path.name}", 3000)
        # The check is quick, run it right away so problems are visible without asking
        self.xml_summation_check()

    def update_title(self):
        if self.export is None:
            self.setWindowTitle("Pol-Tax")
            return
        name = self.export.parser.company_info()["name"] or self.export.xml_path.name
        self.setWindowTitle(f"{name} — Pol-Tax")

    def fill_nav(self):
        self.nav.blockSignals(True)
        self.nav.clear()
        self.section_items = {}
        for section in SECTIONS:
            if section.key not in self.export.sections:
                continue
            item = QTreeWidgetItem(self.nav, [SECTION_TITLES[section.key]])
            item.setIcon(0, self.spreadsheet_icon)
            item.setData(0, Qt.UserRole, ("section", section.key))
            self.section_items[section.key] = item

        attachments = QTreeWidgetItem(self.nav, [f"Załączniki ({len(self.export.attachments)})"])
        attachments.setFlags(attachments.flags() & ~Qt.ItemIsSelectable)
        for path in self.export.attachments:
            child = QTreeWidgetItem(attachments, [path.name])
            child.setIcon(0, self.pdf_icon)
            child.setToolTip(0, f"{path}\nKliknij dwukrotnie, aby otworzyć")
            child.setData(0, Qt.UserRole, ("attachment", str(path)))
        attachments.setExpanded(True)
        self.nav.blockSignals(False)

    def select_section(self, section_key):
        if section_key is None or section_key not in self.section_items:
            return
        self.nav.setCurrentItem(self.section_items[section_key])
        self.show_section(section_key)

    def nav_item_changed(self, current, previous):
        if current is None:
            return
        kind, value = current.data(0, Qt.UserRole) or (None, None)
        if kind == "section":
            self.show_section(value)

    def nav_item_activated(self, item, column=0):
        kind, value = item.data(0, Qt.UserRole) or (None, None)
        if kind == "attachment":
            self.open_path(value)

    def show_nav_menu(self, position):
        item = self.nav.itemAt(position)
        if item is None:
            return
        kind, value = item.data(0, Qt.UserRole) or (None, None)
        menu = QMenu(self)
        if kind == "section":
            menu.addAction(self.excel_icon, "Otwórz w Excelu",
                           lambda: self.open_path(self.export.sections[value][0]))
        elif kind == "attachment":
            menu.addAction(self.pdf_icon, "Otwórz", lambda: self.open_path(value))
            menu.addAction("Pokaż w folderze", lambda: self.open_path(Path(value).parent))
        else:
            return
        menu.exec(self.nav.viewport().mapToGlobal(position))

    def show_section(self, section_key):
        if self.export is None or section_key not in self.export.sections:
            return
        self.current_section = section_key
        _, dataframe, border_list = self.export.sections[section_key]
        is_introduction = section_key == "introduction"
        if is_introduction:
            dataframe = dataframe.rename(columns={"Col1": "", "Col2": "", "Col3": ""})
        else:
            dataframe = dataframe.rename(columns=self.short_year_headers())
        has_years = section_key in YEAR_SECTIONS
        if has_years and self.show_change.isChecked():
            dataframe = with_year_change(dataframe)

        self.current_model = PandasModel(dataframe, self.row_checks(section_key), with_hierarchy=not is_introduction)
        self.current_model.collapsed = self.collapsed_by_section.setdefault(section_key, set())
        self.tableView.setModel(self.current_model)
        self.tableView.setItemDelegate(ColorDelegate(border_list, self.tableView))
        # Names get the room, amounts a fixed width (their headers are long, the full text is in the tooltip)
        header = self.tableView.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        if section_key == "introduction":
            # Group name, field name, value - the row numbers mean nothing here
            self.tableView.verticalHeader().setVisible(False)
            self.tableView.setColumnWidth(0, 300)
            self.tableView.setColumnWidth(1, 260)
            header.setSectionResizeMode(2, QHeaderView.Stretch)
        else:
            self.tableView.verticalHeader().setVisible(True)
            header.setSectionResizeMode(0, QHeaderView.Stretch)
            for column in range(1, self.current_model.columnCount()):
                name = dataframe.columns[column]
                width = 110 if name == CHANGE_PERCENT_COLUMN else 150 if name == CHANGE_COLUMN else 170
                self.tableView.setColumnWidth(column, width)
        self.section_title.setText(SECTION_TITLES[section_key])
        self.show_change.setVisible(has_years)
        self.expand_button.setVisible(not is_introduction)
        self.collapse_button.setVisible(not is_introduction)
        self.apply_row_visibility()
        self.update_actions()

    def apply_row_visibility(self):
        """
        Hides rows inside folded positions, or when searching, rows that don't match
        """
        model = self.current_model
        if model is None:
            return
        text = self.search.text().strip().casefold()
        for row in range(model.rowCount()):
            if text:
                name = str(model.index(row, 0).data(Qt.DisplayRole)).casefold()
                visible = text in name or text in model.key(row).casefold()
            else:
                visible = not any(ancestor in model.collapsed for ancestor in model.ancestors(row))
            self.tableView.setRowHidden(row, not visible)

    def toggle_row(self, row):
        if self.current_model is None or self.search.text():
            return
        self.current_model.toggle(row)
        self.apply_row_visibility()

    def expand_all(self):
        if self.current_model is not None:
            self.current_model.collapsed.clear()
            self.apply_row_visibility()
            self.tableView.viewport().update()

    def collapse_all(self):
        """
        Leaves the total and the main positions (A, B, ...) visible
        """
        model = self.current_model
        if model is None:
            return
        # Every level is folded, unfolding a position then shows just its direct sub-positions
        model.collapsed.clear()
        model.collapsed.update(key for key in model.keys_with_children() if key != "")
        self.apply_row_visibility()
        self.tableView.viewport().update()

    def short_year_headers(self):
        return short_year_headers(self.app_settings.get_keywords(), self.export.parser.company_info()["date_to"])

    def row_checks(self, section_key):
        if self.summation is None:
            return {}
        checks = {}
        for result in self.summation.results:
            if result.section == section_key:
                checks.setdefault(result.index, []).append(result)
        return checks

    # ------------------------------------------------------------------ consistency check

    def xml_summation_check(self):
        if self.export is None:
            self.statusBar().showMessage("Najpierw otwórz plik XML", 1500)
            return
        self.run_in_background("Sprawdzanie sum…", self.show_summation,
                               export_summation, self.export.parser, self.app_settings.get_excel_dir(),
                               self.export.prefix)

    def show_summation(self, summation: SummationResult):
        self.summation = summation
        failures = summation.failures
        checked = plural(len(summation.results), 'regułę', 'reguły', 'reguł')
        if failures:
            found = plural(len(failures), 'niezgodność', 'niezgodności', 'niezgodności')
            self.check_summary.setText(f"Kontrola sum: {found} (sprawdzono {checked})")
            color = FAILED_TEXT_COLOR
        else:
            self.check_summary.setText(f"Kontrola sum: wszystko się zgadza (sprawdzono {checked})")
            color = OK_TEXT_COLOR
        self.check_summary.setStyleSheet(f"font-weight: 600; color: {color.name()};")
        self.fill_check_table()
        self.mark_nav_checks()
        self.check_panel.setVisible(True)
        if self.current_section:
            self.show_section(self.current_section)
        self.statusBar().showMessage(self.check_summary.text(), 5000)

    def mark_nav_checks(self):
        for section_key, item in self.section_items.items():
            results = [r for r in self.summation.results if r.section == section_key]
            failed = sum(1 for r in results if not r.ok)
            title = SECTION_TITLES[section_key]
            if failed:
                item.setText(0, f"{title}  ✗ {failed}")
                item.setForeground(0, FAILED_TEXT_COLOR)
                item.setToolTip(0, f"Niezgodności: {failed}")
            elif results:
                item.setText(0, f"{title}  ✓")
                item.setForeground(0, OK_TEXT_COLOR)
                item.setToolTip(0, "Wszystkie sumy się zgadzają")

    def fill_check_table(self):
        if self.summation is None:
            return
        results = self.summation.failures if self.only_failures.isChecked() else self.summation.results
        self.check_table.setRowCount(len(results))
        for row, result in enumerate(results):
            status = QTableWidgetItem("✓" if result.ok else "✗")
            status.setForeground(OK_TEXT_COLOR if result.ok else FAILED_TEXT_COLOR)
            cells = [status,
                     QTableWidgetItem(SECTION_TITLES[result.section]),
                     QTableWidgetItem(result.index or "razem"),
                     QTableWidgetItem(YEAR_NAMES[result.year]),
                     QTableWidgetItem(format_amount(result.actual)),
                     QTableWidgetItem(format_amount(result.expected)),
                     QTableWidgetItem(format_amount(result.difference)),
                     QTableWidgetItem(result.description)]
            cells[2].setToolTip(str(result.label))
            for column in (4, 5, 6):
                cells[column].setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            cells[7].setToolTip(result.description)
            cells[0].setData(Qt.UserRole, (result.section, result.index))
            for column, cell in enumerate(cells):
                self.check_table.setItem(row, column, cell)

    def check_row_clicked(self, row, column):
        section_key, index = self.check_table.item(row, 0).data(Qt.UserRole)
        self.select_section(section_key)
        table_row = self.current_model.row_of_index(index) if self.current_model else None
        if table_row is not None:
            self.search.clear()
            self.current_model.collapsed.difference_update(self.current_model.ancestors(table_row))
            self.apply_row_visibility()
            self.tableView.selectRow(table_row)
            self.tableView.scrollTo(self.current_model.index(table_row, 0), QAbstractItemView.PositionAtCenter)

    # ------------------------------------------------------------------ other actions

    def export_to_excel(self):
        if self.export is None:
            return
        start_dir = self.last_export_dir or QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        path, _ = QFileDialog.getSaveFileName(self, "Eksportuj do Excela",
                                              str(Path(start_dir) / default_workbook_name(self.export)),
                                              "Skoroszyt Excel (*.xlsx)")
        if not path:
            return
        self.last_export_dir = str(Path(path).parent)
        self.run_in_background("Zapisywanie pliku Excel…", self.export_done, export_workbook, self.export, Path(path),
                               self.app_settings.get_keywords(), self.summation)

    def export_done(self, path):
        answer = QMessageBox.question(self, "Eksport zakończony", f"Zapisano:\n{path}\n\nOtworzyć teraz?")
        if answer == QMessageBox.Yes:
            self.open_path(path)

    def open_current_excel(self):
        if self.export and self.current_section in self.export.sections:
            self.open_path(self.export.sections[self.current_section][0])

    def open_summation_excel(self):
        if self.summation:
            self.open_path(self.summation.path)

    def open_path(self, path):
        logging.debug(f"Opening {path}")
        os.startfile(str(path))

    def open_excel_folder(self):
        self.open_path(self.app_settings.get_excel_dir())

    def show_settings(self):
        dialog = SettingsDialog(self)
        if dialog.exec() == QDialog.Accepted and self.xml_filepath is not None:
            # Labels may have changed, read the file again
            self.setup_excels()

    def show_about(self):
        AboutDialog(self).exec()

    def quit_app(self):
        self.app.quit()

    def load_svg_and_color(self, icon_path, color):
        renderer = QSvgRenderer()
        renderer.load(icon_path)

        pixmap = QPixmap(QSize(64, 64))
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        renderer.render(painter)

        painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
        painter.fillRect(pixmap.rect(), QColor(color))
        painter.end()

        return QIcon(pixmap)
