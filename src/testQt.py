import sys
from PyQt6 import QtWidgets, QtGui, uic
from PyQt6.QtCore import Qt
from document_manager import DocumentManager
from data_manager import DataManager
from table_formatter import TableFormatter
from docx.shared import Inches

class DraggableWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi("./ui/draggable_window.ui", self)
        self.set_background_image()
        self.setAcceptDrops(True)
        self.label.setStyleSheet(self.style_default())

    def style_default(self):
        return """
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                padding: 20px;
                text-align: center;
                color: #777;
            }
        """

    def style_active(self):
        return """
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                padding: 20px;
                text-align: center;
                color: #777;
                background-color: #f0f0f0;
            }
        """
    
    def set_background_image(self):
        pixmap = QtGui.QPixmap("./img/cuteBg.jpg")
        scaled_pixmap = pixmap.scaled(self.size(), Qt.AspectRatioMode.KeepAspectRatioByExpanding, Qt.TransformationMode.SmoothTransformation)
        palette = self.palette()
        palette.setBrush(self.backgroundRole(), QtGui.QBrush(scaled_pixmap))
        self.setPalette(palette)
        self.setAutoFillBackground(True)
        self.update()  # 確保背景圖隨窗口大小改變

    def resizeEvent(self, event):
        self.set_background_image()
        super().resizeEvent(event)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and all(url.fileName().endswith('.xlsx') for url in event.mimeData().urls()):
            event.acceptProposedAction()
            self.label.setStyleSheet(self.style_active())

    def dragLeaveEvent(self, event):
        self.label.setStyleSheet(self.style_default())

    def dropEvent(self, event):
        self.label.setStyleSheet(self.style_default())
        files = [url.toLocalFile() for url in event.mimeData().urls() if url.fileName().endswith('.xlsx')]
        if files:
            self.process_files(files)

    def process_files(self, files):
        for file_path in files:
            cleaned_data = DataManager.load_and_prepare_data(file_path)
            if cleaned_data is not None:
                doc = DocumentManager.create_document()
                column_widths = (Inches(0.8), Inches(0.5), Inches(1.962), Inches(1.24), Inches(0.2), Inches(0.3), Inches(0.2))
                TableFormatter.create_and_format_table(doc, cleaned_data, column_widths)
                output_path = file_path.replace('.xlsx', '.docx')
                DocumentManager.save_document(doc, output_path)
                self.label.setText(f"已生成: {output_path}")

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = DraggableWindow()
    window.show()
    sys.exit(app.exec())
