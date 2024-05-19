import sys
from PyQt6 import QtWidgets
from ui.draggable_window import DraggableWindow

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = DraggableWindow()
    window.show()
    sys.exit(app.exec())
