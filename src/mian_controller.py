import os

from data_manager import DataManager
from table_formatter import TableFormatter
from file_scanner import FileScanner
from document_manager import DocumentManager
from user_interface import UserInterface
from docx.shared import Inches
class MainController:
    def run(self):
        directory = input("請輸入測試案例.xlsx放的路徑: ")
        files = FileScanner.scan_xlsx_files(directory)
        displayed_files = UserInterface.display_files(files)
        if displayed_files:
            selected_file = UserInterface.user_select_file(displayed_files)
            if selected_file:
                print(f"你要產的測試報告：{os.path.basename(selected_file)}")
                cleaned_data = DataManager.load_and_prepare_data(selected_file)
                doc = DocumentManager.create_document()
                column_widths = (Inches(0.8), Inches(0.5), Inches(1.962), Inches(1.24), Inches(0.2), Inches(0.3), Inches(0.2))
                TableFormatter.create_and_format_table(doc, cleaned_data, column_widths)
                output_path = selected_file.split('.')[0] + '.docx'
                DocumentManager.save_document(doc, output_path)
            else:
                print("你沒選 = =")

if __name__ == '__main__':
    controller = MainController()
    controller.run()
