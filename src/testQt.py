import sys
import os
from PyQt6 import QtWidgets, QtGui, QtCore
from PyQt6.QtCore import Qt, pyqtSignal
from document_manager import DocumentManager
from data_manager import DataManager
from table_formatter import TableFormatter
from docx.shared import Inches

def resource_path(relative_path):
    """ 獲取資源的絕對路徑，適用於開發環境和打包後環境 """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

class ClickableLabel(QtWidgets.QLabel):
    clicked = pyqtSignal()

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)

class DraggableWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        # 設置窗口標題和大小
        self.setWindowTitle('測試報告產生器')
        self.setGeometry(100, 100, 400, 300)

        # 創建可點擊的 QLabel
        self.label = ClickableLabel('Drag a file here', self)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setStyleSheet(self.style_default())
        self.label.clicked.connect(self.open_file_dialog)

        # 設置布局
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.label)
        self.setLayout(layout)

        # 允許拖放操作
        self.setAcceptDrops(True)

        self.overlay = QtWidgets.QWidget(self)
        self.overlay.setStyleSheet("background-color: rgba(0, 0, 0, 50%);")
        self.overlay.setGeometry(self.rect())
        self.overlay.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)  # 讓遮罩層透明以允許事件傳遞
        self.overlay.lower()  # 確保遮罩在其他組件下方
        self.overlay.show()

         # 設置背景圖片
        self.set_background_image()
        self.show()

    def style_default(self):
        """ 返回 QLabel 的默認樣式表 """
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
        """ 返回 QLabel 的活動狀態樣式表 """
        return """
            QLabel {
                border: 2px dashed #ffffff;
                border-radius: 10px;
                padding: 20px;
                text-align: center;
                color: #ffffff;
            }
        """

    def set_background_image(self):
        """ 設置窗口背景圖片 """
        pixmap = QtGui.QPixmap(resource_path("./img/cuteBg.jpg"))
        scaled_pixmap = pixmap.scaled(self.size(), Qt.AspectRatioMode.KeepAspectRatioByExpanding, Qt.TransformationMode.SmoothTransformation)
        palette = self.palette()
        palette.setBrush(self.backgroundRole(), QtGui.QBrush(scaled_pixmap))
        self.setPalette(palette)
        self.setAutoFillBackground(True)
        self.update()

    def add_overlay(self):
        """ 添加半透明遮罩 """
        overlay = QtWidgets.QWidget(self)
        overlay.setStyleSheet("background-color: rgba(0, 0, 0, 50%);")
        overlay.setGeometry(self.rect())
        overlay.lower()  # 確保遮罩在其他組件下方
        overlay.show()

    def resizeEvent(self, event):
        """ 當窗口大小改變時，重新設置背景圖片和 QLabel 大小 """
        self.set_background_image()
        self.overlay.setGeometry(self.rect())
        self.label.setStyleSheet(self.style_default())  # 重新設置樣式表確保邊框顯示
        self.label.resize(self.size())
        super().resizeEvent(event)

    def dragEnterEvent(self, event):
        """ 當拖動項目進入窗口時檢查是否接受拖動 """
        if event.mimeData().hasUrls() and all(url.fileName().endswith('.xlsx') for url in event.mimeData().urls()):
            event.acceptProposedAction()
            self.label.setStyleSheet(self.style_active())

    def dragLeaveEvent(self, event):
        """ 當拖動項目離開窗口時恢復 QLabel 樣式 """
        self.label.setStyleSheet(self.style_default())

    def dropEvent(self, event):
        """ 當拖動項目放下時處理文件 """
        self.label.setStyleSheet(self.style_default())
        files = [url.toLocalFile() for url in event.mimeData().urls() if url.fileName().endswith('.xlsx')]
        if files:
            self.process_files(files)

    def open_file_dialog(self):
        """ 打開文件選擇對話框 """
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "選擇文件", "", "Excel Files (*.xlsx);;All Files (*)")
        if files:
            self.process_files(files)

    def process_files(self, files):
        """ 處理拖動進來或選擇的文件 """
        for file_path in files:
            # 加載和準備數據
            cleaned_data = DataManager.load_and_prepare_data(file_path)
            if cleaned_data is not None:
                # 創建文檔
                doc = DocumentManager.create_document()
                # 設置表格列寬
                column_widths = (Inches(0.8), Inches(0.5), Inches(1.962), Inches(1.24), Inches(0.2), Inches(0.3), Inches(0.2))
                # 創建並格式化表格
                TableFormatter.create_and_format_table(doc, cleaned_data, column_widths)
                # 保存文檔
                output_path = file_path.replace('.xlsx', '.docx')
                DocumentManager.save_document(doc, output_path)
                # 更新 QLabel 顯示生成的文件路徑
                self.label.setText(f"已生成: {output_path}")

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = DraggableWindow()
    window.show()
    sys.exit(app.exec())
