from decimal import Decimal

import pandas as pd
from PySide6 import QtCore
from PySide6.QtCore import Qt, QModelIndex
from PySide6.QtGui import QColor

from XML.rules import YEAR_NAMES

CHECK_OK_COLOR = QColor("#dff0d8")
CHECK_FAILED_COLOR = QColor("#f8d7da")


class PandasModel(QtCore.QAbstractTableModel):
    """A model to interface a Qt view with pandas dataframe """

    def __init__(self, dataframe: pd.DataFrame, row_checks=None, parent=None):
        """
        row_checks: dataframe index -> list of CheckResult for that position, colors the row and adds a tooltip
        """
        QtCore.QAbstractTableModel.__init__(self, parent)
        self._dataframe = dataframe
        self._row_checks = row_checks or {}

    def rowCount(self, parent=QModelIndex()) -> int:
        if parent == QModelIndex():
            return len(self._dataframe)
        return 0

    def columnCount(self, parent=QModelIndex()) -> int:
        if parent == QModelIndex():
            return len(self._dataframe.columns)
        return 0

    def row_of_index(self, df_index):
        for row, value in enumerate(self._dataframe.index):
            if value == df_index:
                return row
        return None

    def checks_for_row(self, row):
        return self._row_checks.get(self._dataframe.index[row], [])

    def data(self, index: QModelIndex, role=Qt.ItemDataRole):
        if not index.isValid():
            return None

        value = self._dataframe.iloc[index.row(), index.column()]
        if role == Qt.DisplayRole:
            if value is None or (isinstance(value, float) and value != value):
                return ""
            return str(value)
        if role == Qt.UserRole:
            # Raw value, so the delegate can format amounts and leave texts like NIP alone
            return value
        if role == Qt.TextAlignmentRole:
            if isinstance(value, Decimal):
                return int(Qt.AlignRight | Qt.AlignVCenter)
            return int(Qt.AlignLeft | Qt.AlignVCenter)

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


def format_amount(value):
    return format(value, ',.2f').replace(',', ' ').replace('.', ',')


def check_tooltip(check):
    status = "OK" if check.ok else f"różnica {format_amount(check.difference)}"
    return (f"{YEAR_NAMES[check.year]}: powinno być {check.description} = {format_amount(check.expected)}"
            f" ({status})")
