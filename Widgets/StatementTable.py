from decimal import Decimal

from PySide6.QtCore import QLocale, Qt, Signal
from PySide6.QtGui import QAction, QKeySequence
from PySide6.QtWidgets import QAbstractItemView, QApplication, QTableView

from Model.PandasModel import DEPTH_ROLE, HAS_CHILDREN_ROLE
from Widgets.ColorDelegate import ARROW_WIDTH, INDENT_WIDTH
from XML.presentation import Percent


def clipboard_text(value, decimal_point):
    """
    Cell value as Excel understands it when pasted: plain numbers with the local decimal separator
    """
    if isinstance(value, Percent):
        return f"{value:.1f}%".replace(".", decimal_point)
    if isinstance(value, Decimal):
        return f"{value:.2f}".replace(".", decimal_point)
    if value is None or (isinstance(value, float) and value != value):
        return ""
    return str(value).replace("\t", " ").replace("\n", " ")


class StatementTable(QTableView):
    """
    Table of a statement part: clicking the arrow in the name column folds a position,
    Ctrl+C copies the selected cells (tab separated, ready to paste into Excel)
    """
    toggle_requested = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setContextMenuPolicy(Qt.ActionsContextMenu)
        copy_action = QAction("Kopiuj", self)
        copy_action.setShortcut(QKeySequence.Copy)
        copy_action.setShortcutContext(Qt.WidgetWithChildrenShortcut)
        copy_action.triggered.connect(self.copy_selection)
        self.addAction(copy_action)

    def mousePressEvent(self, event):
        index = self.indexAt(event.position().toPoint())
        if index.isValid() and index.column() == 0 and index.data(HAS_CHILDREN_ROLE):
            arrow_left = self.visualRect(index).left() + 4 + (index.data(DEPTH_ROLE) or 0) * INDENT_WIDTH
            if arrow_left <= event.position().x() <= arrow_left + ARROW_WIDTH:
                self.toggle_requested.emit(index.row())
                return
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event):
        index = self.indexAt(event.position().toPoint())
        if index.isValid() and index.data(HAS_CHILDREN_ROLE):
            self.toggle_requested.emit(index.row())
            return
        super().mouseDoubleClickEvent(event)

    def copy_selection(self):
        model = self.model()
        indexes = [index for index in self.selectedIndexes() if not self.isRowHidden(index.row())]
        if model is None or not indexes:
            return
        rows = sorted({index.row() for index in indexes})
        columns = sorted({index.column() for index in indexes})
        selected = {(index.row(), index.column()) for index in indexes}
        decimal_point = QLocale().decimalPoint()
        whole_rows = len(columns) == model.columnCount()

        lines = []
        if whole_rows and len(rows) > 1:
            lines.append("\t".join(["Pozycja"] + [str(model.headerData(c, Qt.Horizontal, Qt.DisplayRole))
                                                  for c in columns]))
        for row in rows:
            cells = [clipboard_text(model.index(row, column).data(Qt.UserRole), decimal_point)
                     if (row, column) in selected else "" for column in columns]
            if whole_rows:
                cells.insert(0, str(model.headerData(row, Qt.Vertical, Qt.DisplayRole)))
            lines.append("\t".join(cells))
        QApplication.clipboard().setText("\n".join(lines))
