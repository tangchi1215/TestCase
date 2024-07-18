from PyQt6 import QtWidgets
from PyQt6.QtCore import Qt

from src.ui.clickable_label import ClickableLabel


class UIFactory:
    @staticmethod
    def create_button(text, parent=None):
        button = QtWidgets.QPushButton(text, parent)
        return button

    @staticmethod
    def create_label(text, parent=None):
        label = ClickableLabel(text, parent)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        return label

    @staticmethod
    def create_radio_button(text, parent=None, checked=False):
        radio_button = QtWidgets.QRadioButton(text, parent)
        radio_button.setChecked(checked)
        return radio_button
