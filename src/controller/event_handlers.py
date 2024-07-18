import os
import shutil

from PyQt6 import QtWidgets, QtCore

TEMPLATE_XLSX = "D:\\TestCase\\src\\assets\\templates\\template.xlsx"


def print_selection(overwrite_yes_radio, overwrite_no_radio):
    """ 打印覆蓋選項的值 """
    if overwrite_yes_radio.isChecked():
        print("是否覆蓋: 是")
    if overwrite_no_radio.isChecked():
        print("是否覆蓋: 否")


def download_template(parent):
    """ 處理按鈕點擊事件，讓使用者選擇保存 template.xlsx 文件的位置 """
    template_path = TEMPLATE_XLSX

    # 打開文件保存對話框讓使用者選擇保存路徑
    save_path, _ = QtWidgets.QFileDialog.getSaveFileName(parent, "Save Template",
                                                         "template.xlsx",
                                                         "Excel Files (*.xlsx)")

    if save_path:
        try:
            save_path = handle_existing_file(save_path)
            shutil.copyfile(template_path, save_path)
            QtWidgets.QMessageBox.information(parent, 'Success', f'Template saved to {save_path}')
        except Exception as e:
            QtWidgets.QMessageBox.critical(parent, 'Error', f'Failed to save template: {e}')


def on_file_finished(parent, file_path, output_path):
    QtCore.QMetaObject.invokeMethod(parent, "update_ui_on_finished", QtCore.Qt.ConnectionType.QueuedConnection,
                                    QtCore.Q_ARG(str, file_path), QtCore.Q_ARG(str, output_path))


@QtCore.pyqtSlot(str, str)
def update_ui_on_finished(parent):
    parent.completed_files += 1
    parent.check_all_files_completed()


def on_file_error(parent, file_path, error_message):
    """ 處理文件错误事件 """
    QtWidgets.QMessageBox.critical(parent, 'Error', f'處理 {file_path} 時出錯: {error_message}')
    parent.completed_files += 1
    parent.check_all_files_completed()


def show_completion_message(parent):
    """ 顯示完成消息框 """
    if not parent.failed_files:
        QtWidgets.QMessageBox.information(parent, 'Success', '所有文件已成功處理完成！')
    else:
        QtWidgets.QMessageBox.warning(parent, 'Partial Success',
                                      '以下文件處理失敗:\n' + '\n'.join(parent.failed_files))


def handle_existing_file(file_path):
    """ 處理已存在的文件，如果存在則添加數字後綴 """
    base, extension = os.path.splitext(file_path)
    counter = 1
    new_path = file_path

    while os.path.exists(new_path):
        new_path = f"{base}({counter}){extension}"
        counter += 1

    return new_path
