from PyQt6 import QtWidgets
from PyQt6.QtCore import Qt, pyqtSignal


class ClickableLabel(QtWidgets.QLabel):
    clicked = pyqtSignal()

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def mouse_press_event(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)
