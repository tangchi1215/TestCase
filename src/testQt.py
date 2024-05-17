import sys
from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout
from PyQt6.QtCore import Qt

class DraggableWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Draggable Window with File Drop')
        self.setGeometry(100, 100, 400, 300)
        self.setStyleSheet("background-color: lightgray;")

        self.label = QLabel('Drag a file here', self)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        self.setLayout(layout)

        self.setAcceptDrops(True)
        self.show()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if files:
            self.label.setText('\n'.join(files))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = DraggableWindow()
    sys.exit(app.exec())
