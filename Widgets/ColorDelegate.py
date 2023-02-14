from PySide6.QtWidgets import QStyledItemDelegate
from PySide6.QtGui import QColor, QPen, QTextOption
from PySide6.QtCore import Qt
from decimal import Decimal, InvalidOperation


class ColorDelegate(QStyledItemDelegate):
    def __init__(self, color_list, parent=None):
        self.color_list = color_list
        super(ColorDelegate, self).__init__(parent)

    def displayText(self, value, locale):
        try:
            dec = Decimal(value)
            # return locale.format('%.2f', Decimal(value), grouping=True)
            return format(dec, ',').replace(',', ' ').replace('.', ',')
        except InvalidOperation:
            return value

    def paint(self, painter, option, index):
        painter.save()
        row = index.row()
        column = index.column()

        # Set the background color for the first column
        if column == 0:
            painter.fillRect(option.rect, QColor("#d7e4bc"))
        else:
            painter.fillRect(option.rect, option.palette.base())

        # Draw the border around the row
        if self.color_list[row] == 1:
            pen = QPen()
            pen.setColor(Qt.black)
            pen.setWidth(1)
            painter.setPen(pen)
            painter.drawLine(option.rect.bottomLeft(), option.rect.bottomRight())

        # Draw the text in the cell
        display_text = self.displayText(index.data(Qt.DisplayRole), option.locale)
        text_rect = option.rect.translated(0, (option.rect.height() - option.fontMetrics.height()) / 2)
        text_rect.setLeft(text_rect.left() + 4)
        text_option = QTextOption(Qt.AlignLeft | Qt.AlignVCenter)
        text_option.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        # painter.drawText(text_rect, text_option, display_text)
        painter.drawText(text_rect, display_text, text_option)

        painter.restore()
