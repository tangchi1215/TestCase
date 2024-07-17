from functools import partial

from PyQt6 import QtWidgets, QtGui, QtCore
from PyQt6.QtCore import Qt, QThreadPool

from src.service.file_worker import FileWorker
from src.service import event_handlers
from src.ui.clickable_label import ClickableLabel
from src.ui.file_status_widget import FileStatusWidget
from src.utils.resource_path import resource_path

ICON_PATH = "D:\\TestCase\\src\\assets\\img\\cuteIcon.png"
BACKGROUND_PATH = "D:\\TestCase\\src\\assets\\img\\cuteBg.jpg"
LABEL_DEFAULT_STYLE_PATH = "D:\\TestCase\\src\\styles\\label_default.qss"
LABEL_ACTIVE_STYLE_PATH = "D:\\TestCase\\src\\styles\\label_active.qss"
LIST_WIDGET_STYLE_PATH = "D:\\TestCase\\src\\styles\\list_widget.qss"


def load_style(widget, style_path):
    """ 加載 qss 樣式並應用到 widget """
    with open(style_path, "r") as f:
        widget.setStyleSheet(f.read())


class DraggableWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        # 設置窗口標題和大小
        self.setWindowTitle('測試報告產生器')
        self.setGeometry(100, 100, 800, 500)

        # 設置圖標
        icon_path = resource_path(ICON_PATH)
        self.setWindowIcon(QtGui.QIcon(icon_path))

        # 創建可點擊的 QLabel
        self.drag_label = ClickableLabel('Drag a file here', self)
        self.drag_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        load_style(self.drag_label, LABEL_DEFAULT_STYLE_PATH)
        self.drag_label.clicked.connect(self.open_file_dialog)

        # 創建一個按鈕
        self.download_template_btn = QtWidgets.QPushButton('Download Template', self)
        self.download_template_btn.clicked.connect(lambda: event_handlers.download_template(self))

        # 創建一個離開按鈕
        self.exit_btn = QtWidgets.QPushButton('Exit', self)
        self.exit_btn.clicked.connect(self.close_application)

        # 創建一個 QListWidget 顯示轉檔狀態
        self.file_status_list = QtWidgets.QListWidget()
        load_style(self.file_status_list, LIST_WIDGET_STYLE_PATH)

        # 創建自訂區域
        custom_group_box = QtWidgets.QGroupBox("自訂選項")
        custom_layout = QtWidgets.QFormLayout()

        # 覆蓋已存在檔案選項
        self.overwrite_yes_radio = QtWidgets.QRadioButton("是")
        self.overwrite_no_radio = QtWidgets.QRadioButton("否")
        self.overwrite_no_radio.setChecked(True)

        # Print Result
        self.overwrite_yes_radio.toggled.connect(self.printSelection)

        overwrite_layout = QtWidgets.QHBoxLayout()
        overwrite_layout.addWidget(self.overwrite_yes_radio)
        overwrite_layout.addWidget(self.overwrite_no_radio)
        overwrite_widget = QtWidgets.QWidget()
        overwrite_widget.setLayout(overwrite_layout)
        custom_layout.addRow(QtWidgets.QLabel("覆蓋已存在檔案:"), overwrite_widget)

        # 測試編號前綴選項
        # self.prefix_filename_radio = QtWidgets.QRadioButton("依檔名")
        # self.prefix_filename_xlsx_radio = QtWidgets.QRadioButton("依文件內自訂")
        # self.prefix_other_radio = QtWidgets.QRadioButton("其他")
        # self.prefix_other_input = QtWidgets.QLineEdit()
        # self.prefix_filename_radio.setChecked(True)
        # self.prefix_other_input.setEnabled(False)
        # self.prefix_other_radio.toggled.connect(self.prefix_other_input.setEnabled)
        # prefix_layout = QtWidgets.QHBoxLayout()
        # prefix_layout.addWidget(self.prefix_filename_radio)
        # prefix_layout.addWidget(self.prefix_filename_xlsx_radio)
        # prefix_layout.addWidget(self.prefix_other_radio)
        # prefix_layout.addWidget(self.prefix_other_input)
        # prefix_widget = QtWidgets.QWidget()
        # prefix_widget.setLayout(prefix_layout)
        # custom_layout.addRow(QtWidgets.QLabel("測試編號前綴:"), prefix_widget)

        # 測試日期選項
        # self.date_today_radio = QtWidgets.QRadioButton("今天")
        # self.date_other_radio = QtWidgets.QRadioButton("其他")
        # self.date_other_input = QtWidgets.QDateEdit()
        # self.date_today_radio.setChecked(True)
        # self.date_other_input.setEnabled(False)
        # self.date_other_radio.toggled.connect(self.date_other_input.setEnabled)
        # date_layout = QtWidgets.QHBoxLayout()
        # date_layout.addWidget(self.date_today_radio)
        # date_layout.addWidget(self.date_other_radio)
        # date_layout.addWidget(self.date_other_input)
        # date_widget = QtWidgets.QWidget()
        # date_widget.setLayout(date_layout)
        # custom_layout.addRow(QtWidgets.QLabel("測試日期:"), date_widget)
        #
        custom_group_box.setLayout(custom_layout)

        # 設置左側布局
        left_layout = QtWidgets.QVBoxLayout()
        left_layout.addWidget(self.drag_label)
        left_layout.addWidget(custom_group_box)
        left_layout.addWidget(self.download_template_btn)
        left_layout.addWidget(self.exit_btn)  # 添加離開按鈕
        left_layout.setStretch(0, 1)  # 讓 QLabel 占據更多空間
        left_layout.setStretch(1, 0)  # 讓自訂選項區占據最小空間
        left_layout.setStretch(2, 0)  # 讓按鈕占據最小空間
        left_layout.setStretch(3, 0)  # 讓按鈕占據最小空間
        left_widget = QtWidgets.QWidget()
        left_widget.setLayout(left_layout)

        # 創建一個 QSplitter 將左側和右側分割
        splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(self.file_status_list)
        splitter.setStretchFactor(0, 8)
        splitter.setStretchFactor(1, 4)

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

    def printSelection(self):
        if self.overwrite_yes_radio.isChecked():
            print("是否覆蓋: 是")
        if self.overwrite_no_radio.isChecked():
            print("是否覆蓋: 否")

    def set_background_image(self):
        """ 設置窗口背景圖片 """
        pixmap = QtGui.QPixmap(BACKGROUND_PATH)
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
        load_style(self.drag_label, LABEL_DEFAULT_STYLE_PATH)  # 重新設置樣式表確保邊框顯示
        self.drag_label.resize(self.size())
        super().resizeEvent(event)

    def dragEnterEvent(self, event):
        """ 當拖動項目進入窗口時檢查是否接受拖動 """
        if event.mimeData().hasUrls() and all(url.fileName().endswith('.xlsx') for url in event.mimeData().urls()):
            event.acceptProposedAction()
            load_style(self.drag_label, LABEL_ACTIVE_STYLE_PATH)

    def dragLeaveEvent(self, event):
        """ 當拖動項目離開窗口時恢復 QLabel 樣式 """
        load_style(self.drag_label, LABEL_DEFAULT_STYLE_PATH)

    def dropEvent(self, event):
        """ 當拖動項目放下時處理文件 """
        load_style(self.drag_label, LABEL_DEFAULT_STYLE_PATH)
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
            worker = FileWorker(file_path, self.overwrite_yes_radio.isChecked())
            worker.signals.progress.connect(file_status_widget.increment_progress)
            worker.signals.finished.connect(partial(event_handlers.on_file_finished, self, file_path))
            worker.signals.error.connect(partial(event_handlers.on_file_error, self, file_path))
            self.thread_pool.start(worker)

    @QtCore.pyqtSlot(str, str)
    def update_ui_on_finished(self, file_path, output_path):
        event_handlers.update_ui_on_finished(self, file_path, output_path)

    def check_all_files_completed(self):
        """ 檢查是否所有文件都已完成處理 """
        if self.completed_files == self.total_files:
            QtCore.QTimer.singleShot(1000, lambda: event_handlers.show_completion_message(self))

    def close_application(self):
        """ 關閉應用程序 """
        self.close()
