from PyQt6 import QtWidgets, QtGui, QtCore
from PyQt6.QtCore import Qt, QThreadPool, QDate

from src.controller import event_handlers
from src.utils.resource_path import resource_path
from src.utils.style_loader import load_style
from src.utils.ui_factory import UIFactory


ICON_PATH = "src/assets/img/cuteIcon.png"
BACKGROUND_PATH = "src/assets/img/cuteBg.jpg"
LABEL_DEFAULT_STYLE_PATH = "src/styles/label_default.qss"
LABEL_ACTIVE_STYLE_PATH = "src/styles/label_active.qss"
LIST_WIDGET_STYLE_PATH = "src/styles/list_widget.qss"
GOODBYE_GIF_PATH = "src/assets/img/goodbye_gif.GIF"
GOODBYE_SOUND_PATH = "src/assets/sound/goodbye_sound.mp3"

# ICON_PATH = "D:\\TestCase\\src\\assets\\img\\cuteIcon.png"
# BACKGROUND_PATH = "D:\\TestCase\\src\\assets\\img\\cuteBg.jpg"
# LABEL_DEFAULT_STYLE_PATH = "D:\\TestCase\\src\\styles\\label_default.qss"
# LABEL_ACTIVE_STYLE_PATH = "D:\\TestCase\\src\\styles\\label_active.qss"
# LIST_WIDGET_STYLE_PATH = "D:\\TestCase\\src\\styles\\list_widget.qss"


class DraggableWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.drag_label = None
        self.download_template_btn = None
        self.exit_btn = None
        self.file_status_list = None
        self.overlay = None
        self.overwrite_yes_radio = None
        self.overwrite_no_radio = None
        self.seqNo_by_filename_radio = None
        self.seqNo_by_excel_radio = None
        self.seqNo_by_custom_radio = None
        self.seqNo_custom_input = None
        self.date_today_radio = None
        self.date_custom_radio = None
        self.date_custom_input = None
        self.thread_pool = QThreadPool()
        self.total_files = 0
        self.completed_files = 0
        self.failed_files = []

        self.init_ui()
        self.init_logic()

    def init_ui(self):
        # 設置窗口標題和大小
        self.setWindowTitle('測試報告產生器')
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        size = self.geometry()
        self.setGeometry((screen.width() - size.width()) // 2, (screen.height() - size.height()) // 2
                         , 800, 500)

        # 設置圖標
        icon_path = resource_path(ICON_PATH)
        self.setWindowIcon(QtGui.QIcon(icon_path))

        # 創建可點擊的 QLabel
        self.drag_label = UIFactory.create_label('Drag a file here', self)
        self.drag_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        load_style(self.drag_label, resource_path(LABEL_DEFAULT_STYLE_PATH))
        self.drag_label.clicked.connect(self.open_file_dialog)

        # 創建一個按鈕
        self.download_template_btn = UIFactory.create_button('Download Template', self)

        # 創建一個離開按鈕
        self.exit_btn = UIFactory.create_button('Exit', self)

        # 創建一個 QListWidget 顯示轉檔狀態
        self.file_status_list = QtWidgets.QListWidget()
        load_style(self.file_status_list, resource_path(LIST_WIDGET_STYLE_PATH))

        custom_group_box = self.create_custom_group_box()

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

        # 初始化文件計數器和狀態列表
        self.total_files = 0
        self.completed_files = 0
        self.failed_files = []

    def init_logic(self):
        self.download_template_btn.clicked.connect(lambda: event_handlers.download_template(self))
        self.exit_btn.clicked.connect(self.show_goodbye_window)
        self.seqNo_by_custom_radio.toggled.connect(self.seqNo_custom_input.setEnabled)
        self.date_custom_radio.toggled.connect(self.date_custom_input.setEnabled)
        # 初始化 QThreadPool
        self.thread_pool = QThreadPool()

    def create_custom_group_box(self):
        # 創建自訂區域
        custom_group_box = QtWidgets.QGroupBox("自訂選項")
        custom_layout = QtWidgets.QFormLayout()

        # 覆蓋已存在檔案選項
        self.overwrite_yes_radio = UIFactory.create_radio_button("是", self)
        self.overwrite_no_radio = UIFactory.create_radio_button("否", self, checked=True)

        overwrite_layout = QtWidgets.QHBoxLayout()
        overwrite_layout.addWidget(self.overwrite_yes_radio)
        overwrite_layout.addWidget(self.overwrite_no_radio)
        overwrite_widget = QtWidgets.QWidget()
        overwrite_widget.setLayout(overwrite_layout)
        custom_layout.addRow(QtWidgets.QLabel("覆蓋已存在檔案:"), overwrite_widget)

        # 測試編號前綴選項
        self.seqNo_by_filename_radio = UIFactory.create_radio_button("檔名_流水號", self, checked=True)
        self.seqNo_by_excel_radio = UIFactory.create_radio_button("讀取excel測試編號", self)
        self.seqNo_by_custom_radio = UIFactory.create_radio_button("自訂前綴_流水號", self)
        self.seqNo_custom_input = QtWidgets.QLineEdit(self)
        self.seqNo_custom_input.setEnabled(False)

        seq_no_layout = QtWidgets.QHBoxLayout()
        seq_no_layout.addWidget(self.seqNo_by_filename_radio)
        seq_no_layout.addWidget(self.seqNo_by_excel_radio)
        seq_no_layout.addWidget(self.seqNo_by_custom_radio)
        seq_no_layout.addWidget(self.seqNo_custom_input)
        seq_no_widget = QtWidgets.QWidget()
        seq_no_widget.setLayout(seq_no_layout)
        custom_layout.addRow(QtWidgets.QLabel("測試編號:"), seq_no_widget)

        # 測試日期選項
        self.date_today_radio = UIFactory.create_radio_button("今天", self, checked=True)
        self.date_custom_radio = UIFactory.create_radio_button("自訂", self)
        self.date_custom_input = QtWidgets.QDateEdit(self)
        self.date_custom_input.setCalendarPopup(True)
        self.date_custom_input.setDate(QDate.currentDate())
        self.date_custom_input.setEnabled(False)

        date_layout = QtWidgets.QHBoxLayout()
        date_layout.addWidget(self.date_today_radio)
        date_layout.addWidget(self.date_custom_radio)
        date_layout.addWidget(self.date_custom_input)
        date_widget = QtWidgets.QWidget()
        date_widget.setLayout(date_layout)
        custom_layout.addRow(QtWidgets.QLabel("測試日期:"), date_widget)

        custom_group_box.setLayout(custom_layout)
        return custom_group_box

    def set_background_image(self):
        """ 設置窗口背景圖片 """
        pixmap = QtGui.QPixmap(resource_path(BACKGROUND_PATH))
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
        load_style(self.drag_label, resource_path(LABEL_DEFAULT_STYLE_PATH))  # 重新設置樣式表確保邊框顯示
        self.drag_label.resize(self.size())
        super().resizeEvent(event)

    def dragEnterEvent(self, event):
        """ 當拖動項目進入窗口時檢查是否接受拖動 """
        if event.mimeData().hasUrls() and all(url.fileName().endswith('.xlsx') for url in event.mimeData().urls()):
            event.acceptProposedAction()
            load_style(self.drag_label, resource_path(LABEL_ACTIVE_STYLE_PATH))

    def dragLeaveEvent(self, event):
        """ 當拖動項目離開窗口時恢復 QLabel 樣式 """
        load_style(self.drag_label, resource_path(LABEL_DEFAULT_STYLE_PATH))

    def dropEvent(self, event):
        """ 當拖動項目放下時處理文件 """
        load_style(self.drag_label, resource_path(LABEL_DEFAULT_STYLE_PATH))
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

    def check_all_files_completed(self):
        """ 檢查是否所有文件都已完成處理 """
        if self.completed_files == self.total_files:
            QtCore.QTimer.singleShot(1000, lambda: event_handlers.show_completion_message(self))

    def show_goodbye_window(self):
        from src.ui.goodbye_window import GoodbyeWindow
        gif_path = resource_path(GOODBYE_GIF_PATH)
        sound_path = resource_path(GOODBYE_SOUND_PATH)
        goodbye_window = GoodbyeWindow(gif_path, sound_path)
        goodbye_window.finished.connect(self.close_application)
        goodbye_window.exec()

    def close_application(self):
        """ 關閉應用程序 """
        self.close()

    @QtCore.pyqtSlot(str, str)
    def update_ui_on_finished(self):
        event_handlers.update_ui_on_finished(self)

    def process_files(self, files):
        from src.service.worker.file_worker import FileWorker
        import datetime
        from functools import partial
        from src.ui.file_status_widget import FileStatusWidget

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

            # 獲取選擇的日期
            if self.date_custom_radio.isChecked():
                selected_date = self.date_custom_input.date().toPyDate()
            else:
                selected_date = QDate.currentDate().toPyDate()

            # 初始化測試編號選項
            seq_no_option = "filename"  # 默認值

            # 獲取測試編號選項
            if self.seqNo_by_filename_radio.isChecked():
                seq_no_option = "filename"
            elif self.seqNo_by_excel_radio.isChecked():
                seq_no_option = "excel"
            elif self.seqNo_by_custom_radio.isChecked():
                seq_no_option = "custom"
            custom_prefix = self.seqNo_custom_input.text()

            # 將日期轉換為 datetime.datetime 物件
            selected_date = datetime.datetime.combine(selected_date, datetime.datetime.min.time())

            worker = FileWorker(file_path,
                                self.overwrite_yes_radio.isChecked(),
                                selected_date,
                                seq_no_option,
                                custom_prefix)

            worker.signals.progress.connect(file_status_widget.increment_progress)
            worker.signals.finished.connect(partial(event_handlers.on_file_finished, self, file_path))
            worker.signals.error.connect(partial(event_handlers.on_file_error, self, file_path))
            self.thread_pool.start(worker)
