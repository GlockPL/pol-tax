from decimal import Decimal

import pandas as pd
from PySide6 import QtCore
from PySide6.QtCore import Qt, QModelIndex
from PySide6.QtGui import QColor

from XML.presentation import Percent, format_amount, format_percent, hierarchy
from XML.rules import YEAR_NAMES

CHECK_OK_COLOR = QColor("#dff0d8")
CHECK_FAILED_COLOR = QColor("#f8d7da")

# Extra data roles for the name column
DEPTH_ROLE = Qt.UserRole + 1
HAS_CHILDREN_ROLE = Qt.UserRole + 2
EXPANDED_ROLE = Qt.UserRole + 3


class PandasModel(QtCore.QAbstractTableModel):
    """A model to interface a Qt view with pandas dataframe """

    def __init__(self, dataframe: pd.DataFrame, row_checks=None, with_hierarchy=True, parent=None):
        """
        row_checks: dataframe index -> list of CheckResult for that position, colors the row and adds a tooltip
        with_hierarchy: positions are nested by their keys (A > A.I > A.I.1) and can be folded
        """
        QtCore.QAbstractTableModel.__init__(self, parent)
        self._dataframe = dataframe
        self._row_checks = row_checks or {}
        self._levels = hierarchy(dataframe.index) if with_hierarchy else {}
        self.collapsed = set()

    def rowCount(self, parent=QModelIndex()) -> int:
        if parent == QModelIndex():
            return len(self._dataframe)
        return 0

    def columnCount(self, parent=QModelIndex()) -> int:
        if parent == QModelIndex():
            return len(self._dataframe.columns)
        return 0

    def key(self, row):
        return str(self._dataframe.index[row])

    def row_of_index(self, df_index):
        for row, value in enumerate(self._dataframe.index):
            if value == df_index:
                return row
        return None

    def level(self, row):
        """
        (parent key, depth, has children) of the position in the row
        """
        return self._levels.get(self.key(row), (None, 0, False))

    def ancestors(self, row):
        parent = self.level(row)[0]
        while parent is not None:
            yield parent
            parent = self._levels.get(parent, (None,))[0]

    def keys_with_children(self):
        return {key for key, (_, _, has_children) in self._levels.items() if has_children}

    def checks_for_row(self, row):
        return self._row_checks.get(self._dataframe.index[row], [])

    def toggle(self, row):
        key = self.key(row)
        if key in self.collapsed:
            self.collapsed.remove(key)
        else:
            self.collapsed.add(key)
        self.dataChanged.emit(self.index(row, 0), self.index(row, 0))

    def data(self, index: QModelIndex, role=Qt.ItemDataRole):
        if not index.isValid():
            return None

        value = self._dataframe.iloc[index.row(), index.column()]
        if role == Qt.DisplayRole:
            if value is None or (isinstance(value, float) and value != value):
                return ""
            if isinstance(value, Percent):
                return format_percent(value)
            if isinstance(value, Decimal):
                return format_amount(value)
            return str(value)
        if role == Qt.UserRole:
            # Raw value, so the delegate can format amounts and leave texts like NIP alone
            return value
        if role == Qt.TextAlignmentRole:
            if isinstance(value, Decimal):
                return int(Qt.AlignRight | Qt.AlignVCenter)
            return int(Qt.AlignLeft | Qt.AlignVCenter)
        if role in (DEPTH_ROLE, HAS_CHILDREN_ROLE, EXPANDED_ROLE):
            if index.column() != 0 or not self._levels:
                return None
            _, depth, has_children = self.level(index.row())
            if role == DEPTH_ROLE:
                return depth
            if role == HAS_CHILDREN_ROLE:
                return has_children
            return self.key(index.row()) not in self.collapsed

        checks = self.checks_for_row(index.row())
        failed = [check for check in checks if not check.ok]
        if role == Qt.BackgroundRole and len(checks) > 0:
            return CHECK_FAILED_COLOR if failed else CHECK_OK_COLOR
        if role == Qt.ToolTipRole:
            # Full text of long names (they are elided in the table) and the check results of the row
            lines = [value] if isinstance(value, str) and value else []
            lines += [check_tooltip(check) for check in (failed or checks)]
            return "\n".join(lines) or None

        return None

    def headerData(
            self, section: int, orientation: Qt.Orientation, role: Qt.ItemDataRole
    ):
        if role in (Qt.DisplayRole, Qt.ToolTipRole):
            if orientation == Qt.Horizontal:
                return str(self._dataframe.columns[section])

            if orientation == Qt.Vertical:
                return str(self._dataframe.index[section])

        return None


def check_tooltip(check):
    status = "OK" if check.ok else f"różnica {format_amount(check.difference)}"
    return (f"{YEAR_NAMES[check.year]}: powinno być {check.description} = {format_amount(check.expected)}"
            f" ({status})")
