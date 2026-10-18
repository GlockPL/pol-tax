from decimal import Decimal

from PySide6.QtCore import QPointF, Qt
from PySide6.QtGui import QColor, QPen, QPolygonF
from PySide6.QtWidgets import QStyle, QStyledItemDelegate

from Model.PandasModel import DEPTH_ROLE, EXPANDED_ROLE, HAS_CHILDREN_ROLE

# Neutral, so green only ever means "check passed"
NAME_COLUMN_COLOR = QColor("#f3f4f6")
INDENT_WIDTH = 16
ARROW_WIDTH = 14


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

        if selected:
            text_color = option.palette.highlightedText().color()
        elif check_color is not None or column == 0:
            # Fixed light backgrounds need dark text, also in dark mode
            text_color = QColor(Qt.black)
        else:
            text_color = option.palette.text().color()

        text_rect = option.rect.adjusted(4, 0, -6, 0)
        has_children = index.data(HAS_CHILDREN_ROLE)
        depth = index.data(DEPTH_ROLE)
        if column == 0 and depth is not None:
            # Nested positions are indented, the ones with sub-positions get a fold arrow
            text_rect.setLeft(text_rect.left() + depth * INDENT_WIDTH)
            if has_children:
                self.draw_arrow(painter, text_rect, index.data(EXPANDED_ROLE), text_color)
            text_rect.setLeft(text_rect.left() + ARROW_WIDTH)
            font = option.font
            font.setBold(bool(has_children))
            painter.setFont(font)

        value = index.data(Qt.UserRole)
        text = index.data(Qt.DisplayRole) or ""
        alignment = Qt.AlignRight | Qt.AlignVCenter if isinstance(value, Decimal) else Qt.AlignLeft | Qt.AlignVCenter

        painter.setPen(text_color)
        # One line per row, long names are elided (the full name is in the tooltip)
        text = painter.fontMetrics().elidedText(text, Qt.ElideRight, text_rect.width())
        painter.drawText(text_rect, alignment, text)

        painter.restore()

    @staticmethod
    def draw_arrow(painter, rect, expanded, color):
        center_x = rect.left() + ARROW_WIDTH / 2 - 1
        center_y = rect.center().y()
        size = 4
        if expanded:
            points = [QPointF(center_x - size, center_y - size / 2), QPointF(center_x + size, center_y - size / 2),
                      QPointF(center_x, center_y + size / 2 + 1)]
        else:
            points = [QPointF(center_x - size / 2, center_y - size), QPointF(center_x - size / 2, center_y + size),
                      QPointF(center_x + size / 2 + 1, center_y)]
        painter.setRenderHint(painter.RenderHint.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(color)
        painter.drawPolygon(QPolygonF(points))
