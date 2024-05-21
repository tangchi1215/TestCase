import sys
import shutil

from PyQt6 import QtWidgets, QtGui, QtCore
from PyQt6.QtCore import Qt,  QThreadPool
from service.file_worker import FileWorker

from ui.clickable_label import ClickableLabel
from ui.file_status_widget import FileStatusWidget

from utils.resource_path import resource_path


def draggable_style_default(widget):
    """ 返回 QLabel 的默認樣式表 """
    widget.setStyleSheet(
        """
        QLabel {
            border: 2px dashed #aaa;
            border-radius: 10px;
            padding: 20px;
            text-align: center;
            color: #777;
        }
    """
    )


def draggable_style_active(widget):
    """ 返回 QLabel 的活動狀態樣式表 """
    widget.setStyleSheet(
        """
        QLabel {
            border: 2px dashed #ffffff;
            border-radius: 10px;
            padding: 20px;
            text-align: center;
            color: #ffffff;
        }
    """
    )


def file_status_list_style(widget):
    """ 返回 QListWidget 的樣式表 """
    widget.setStyleSheet(
        """
        QListWidget {
            background-color: rgba(0, 0, 0, 0);;
            border-radius: 10px;
            padding: 10px;
        }
        QListWidget::item {
            color: #333;
            padding: 5px;
        }
        QListWidget::item:selected {
            background-color: #6c757d;
            color: white;
        }
    """
    )


class DraggableWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        # 設置窗口標題和大小
        self.setWindowTitle('測試報告產生器')
        self.setGeometry(100, 100, 700, 400)

        # 設置圖標
        icon_path = resource_path("src/assets/img/cuteIcon.png")
        # icon_path = "src/assets/img/cuteIcon.png"
        self.setWindowIcon(QtGui.QIcon(icon_path))

        # 創建可點擊的 QLabel
        self.label = ClickableLabel('Drag a file here', self)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        draggable_style_default(self.label)
        self.label.clicked.connect(self.open_file_dialog)

        # 創建一個按鈕
        self.button = QtWidgets.QPushButton('Download Template', self)
        self.button.clicked.connect(self.on_button_click)

        # 創建一個 QListWidget 顯示轉檔狀態
        self.file_status_list = QtWidgets.QListWidget()
        file_status_list_style(self.file_status_list)

        # 設置左側布局
        left_layout = QtWidgets.QVBoxLayout()
        left_layout.addWidget(self.label)
        left_layout.addWidget(self.button)
        left_layout.setStretch(0, 1)  # 讓 QLabel 占據更多空間
        left_layout.setStretch(1, 0)  # 讓按鈕占據最小空間
        left_widget = QtWidgets.QWidget()
        left_widget.setLayout(left_layout)

        # 創建一個 QSplitter 將左側和右側分割
        splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(self.file_status_list)
        splitter.setStretchFactor(0, 10)
        splitter.setStretchFactor(1, 2)

        # 設置主布局
        main_layout = QtWidgets.QVBoxLayout()
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

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

        # 初始化 QThreadPool
        self.thread_pool = QThreadPool()

        # 初始化文件計數器和狀態列表
        self.total_files = 0
        self.completed_files = 0
        self.failed_files = []

    def set_background_image(self):
        """ 設置窗口背景圖片 """
        # pixmap = QtGui.QPixmap("src/assets/img/cuteBg.jpg")
        pixmap = QtGui.QPixmap(resource_path("src/assets/img/cuteBg.jpg"))
        scaled_pixmap = pixmap.scaled(self.size(), Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                                      Qt.TransformationMode.SmoothTransformation)
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
        draggable_style_default(self.label)  # 重新設置樣式表確保邊框顯示
        self.label.resize(self.size())
        super().resizeEvent(event)

    def dragEnterEvent(self, event):
        """ 當拖動項目進入窗口時檢查是否接受拖動 """
        if event.mimeData().hasUrls() and all(url.fileName().endswith('.xlsx') for url in event.mimeData().urls()):
            event.acceptProposedAction()
            draggable_style_active(self.label)

    def dragLeaveEvent(self, event):
        """ 當拖動項目離開窗口時恢復 QLabel 樣式 """
        draggable_style_default(self.label)

    def dropEvent(self, event):
        """ 當拖動項目放下時處理文件 """
        draggable_style_default(self.label)
        files = [url.toLocalFile() for url in event.mimeData().urls() if url.fileName().endswith('.xlsx')]
        if files:
            self.process_files(files)

    def open_file_dialog(self):
        """ 打開文件選擇對話框 """
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(self,
                                                          "選擇文件", "",
                                                          "Excel Files (*.xlsx);;All Files (*)")
        if files:
            self.process_files(files)

    def process_files(self, files):
        """ 處理拖動進來或選擇的文件 """
        self.total_files = len(files)
        self.completed_files = 0
        self.failed_files = []

        for file_path in files:
            file_status_widget = FileStatusWidget(file_path)
            item = QtWidgets.QListWidgetItem()
            item.setSizeHint(file_status_widget.sizeHint())
            self.file_status_list.addItem(item)
            self.file_status_list.setItemWidget(item, file_status_widget)
            QtWidgets.QApplication.processEvents()  # 更新 UI

            # 創建並運行 FileWorker
            worker = FileWorker(file_path)
            worker.signals.progress.connect(file_status_widget.increment_progress)
            worker.signals.finished.connect(self.on_file_finished)
            worker.signals.error.connect(self.on_file_error)
            self.thread_pool.start(worker)

    def on_file_finished(self, file_path, output_path):
        QtCore.QMetaObject.invokeMethod(self, "update_ui_on_finished", QtCore.Qt.ConnectionType.QueuedConnection,
                                        QtCore.Q_ARG(str, file_path), QtCore.Q_ARG(str, output_path))

    @QtCore.pyqtSlot(str, str)
    def update_ui_on_finished(self, file_path, output_path):
        self.completed_files += 1
        self.check_all_files_completed()

    def on_file_error(self, file_path, error_message):
        """ 處理文件错误事件 """
        QtWidgets.QMessageBox.critical(self, 'Error', f'處理 {file_path} 時出錯: {error_message}')
        self.completed_files += 1
        self.check_all_files_completed()

    def check_all_files_completed(self):
        """ 檢查是否所有文件都已完成處理 """
        if self.completed_files == self.total_files:
            QtCore.QTimer.singleShot(1000, self.show_completion_message)

    def show_completion_message(self):
        """ 顯示完成消息框 """
        if not self.failed_files:
            QtWidgets.QMessageBox.information(self, 'Success', '所有文件已成功處理完成！')
        else:
            QtWidgets.QMessageBox.warning(self, 'Partial Success',
                                          '以下文件處理失敗:\n' + '\n'.join(self.failed_files))

    def on_button_click(self):
        """ 處理按鈕點擊事件，讓使用者選擇保存 template.xlsx 文件的位置 """
        # template_path = "src/assets/templates/template.xlsx"
        template_path = resource_path("src/assets/templates/template.xlsx")

        # 打開文件保存對話框讓使用者選擇保存路徑
        save_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Template",
                                                             "template.xlsx",
                                                             "Excel Files (*.xlsx)")

        if save_path:
            try:
                shutil.copyfile(template_path, save_path)
                QtWidgets.QMessageBox.information(self, 'Success', f'Template saved to {save_path}')
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, 'Error', f'Failed to save template: {e}')


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = DraggableWindow()
    window.show()
    sys.exit(app.exec())
