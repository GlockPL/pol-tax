from decimal import Decimal

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPen
from PySide6.QtWidgets import QStyle, QStyledItemDelegate

from Model.PandasModel import format_amount

# Neutral, so green only ever means "check passed"
NAME_COLUMN_COLOR = QColor("#f3f4f6")


class ColorDelegate(QStyledItemDelegate):
    def __init__(self, color_list, parent=None):
        self.color_list = color_list
        super(ColorDelegate, self).__init__(parent)

    def paint(self, painter, option, index):
        painter.save()
        row = index.row()
        column = index.column()

        check_color = index.data(Qt.BackgroundRole)
        selected = option.state & QStyle.State_Selected
        if selected:
            painter.fillRect(option.rect, option.palette.highlight())
        elif check_color is not None:
            painter.fillRect(option.rect, check_color)
        elif column == 0:
            painter.fillRect(option.rect, NAME_COLUMN_COLOR)
        else:
            painter.fillRect(option.rect, option.palette.base())

        # Draw the border under header rows
        if row < len(self.color_list) and self.color_list[row] == 1:
            pen = QPen()
            pen.setColor(Qt.black)
            pen.setWidth(1)
            painter.setPen(pen)
            painter.drawLine(option.rect.bottomLeft(), option.rect.bottomRight())

        value = index.data(Qt.UserRole)
        if isinstance(value, Decimal):
            text = format_amount(value)
            alignment = Qt.AlignRight | Qt.AlignVCenter
        else:
            text = index.data(Qt.DisplayRole) or ""
            alignment = Qt.AlignLeft | Qt.AlignVCenter

        if selected:
            painter.setPen(option.palette.highlightedText().color())
        elif check_color is not None or column == 0:
            # Fixed light backgrounds need dark text, also in dark mode
            painter.setPen(Qt.black)
        else:
            painter.setPen(option.palette.text().color())
        text_rect = option.rect.adjusted(4, 0, -6, 0)
        # One line per row, long names are elided (the full name is in the tooltip)
        text = option.fontMetrics.elidedText(text, Qt.ElideRight, text_rect.width())
        painter.drawText(text_rect, alignment, text)

        painter.restore()
