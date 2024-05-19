import os
from PyQt6 import QtWidgets


class FileStatusWidget(QtWidgets.QWidget):
    def __init__(self, file_name, parent=None):
        super(FileStatusWidget, self).__init__(parent)

        # 垂直布局
        main_layout = QtWidgets.QVBoxLayout()

        # 水平布局(包含檔案名稱, 百分比)
        top_layout = QtWidgets.QHBoxLayout()

        # 檔案名稱
        self.file_name_label = QtWidgets.QLabel(os.path.basename(file_name))

        # 百分比
        self.percentage_label = QtWidgets.QLabel("0%")

        top_layout.addWidget(self.file_name_label)
        top_layout.addStretch()  # 添加伸缩空间，使百分比标签靠右
        top_layout.addWidget(self.percentage_label)

        # 進度條
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)

        # 添加组件到垂直布局
        main_layout.addLayout(top_layout)
        main_layout.addWidget(self.progress_bar)

        self.setLayout(main_layout)

    def update_progress(self, value):
        """ 更新進度條跟百分比 """
        self.progress_bar.setValue(value)
        self.percentage_label.setText(f"{value}%")
