from PyQt6.QtCore import QRunnable, pyqtSignal, QObject
from service.data_manager import DataManager
from service.document_manager import DocumentManager
from service.table_formatter import TableFormatter
from docx.shared import Inches


class WorkerSignals(QObject):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str, str)
    error = pyqtSignal(str, str)


class FileWorker(QRunnable):
    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path
        self.signals = WorkerSignals()

    def run(self):
        try:
            self.signals.progress.emit(0) # 更新進度
            cleaned_data = DataManager.load_and_prepare_data(self.file_path)
            self.signals.progress.emit(20)  # 更新進度

            if cleaned_data is None:
                raise ValueError("Failed to process data")

            doc = DocumentManager.create_document()
            column_widths = (
                Inches(0.8), Inches(0.5), Inches(1.962), Inches(1.24), Inches(0.2), Inches(0.3), Inches(0.2))
            TableFormatter.create_and_format_table(doc, cleaned_data, column_widths)
            self.signals.progress.emit(70)  # 更新進度

            output_path = self.file_path.replace('.xlsx', '.docx')
            DocumentManager.save_document(doc, output_path)
            self.signals.progress.emit(100) # 更新進度
            self.signals.finished.emit(self.file_path, output_path)
        except Exception as e:
            self.signals.error.emit(self.file_path, str(e))
