import sys
from PyQt6 import QtWidgets, uic
from PyQt6.QtCore import Qt

class DraggableWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi("./ui/draggable_window.ui", self)

        # 設置初始樣式
        self.label.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                padding: 20px;
                text-align: center;
                color: #777;
            }
        """)

        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.label.setStyleSheet("""
                QLabel {
                    border: 2px dashed #aaa;
                    border-radius: 10px;
                    padding: 20px;
                    text-align: center;
                    color: #777;
                    background-color: #f0f0f0;
                }
            """)

    def dragLeaveEvent(self, event):
        self.label.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                padding: 20px;
                text-align: center;
                color: #777;
            }
        """)

    def dropEvent(self, event):
        self.label.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                padding: 20px;
                text-align: center;
                color: #777;
            }
        """)
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if files:
            self.label.setText('\n'.join(files))

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = DraggableWindow()
    window.show()
    sys.exit(app.exec())
