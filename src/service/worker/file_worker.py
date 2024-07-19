from PyQt6.QtCore import QRunnable
from src.service.data_manager_service import DataManagerService
from src.service.document_manager_service import DocumentManagerService
from src.service.table_formatter_service import TableFormatterService
from src.controller import event_handlers
from src.service.worker.worker_signals import WorkerSignals
from docx.shared import Inches
from src.domain.file_metadata import FileMetadata


class FileWorker(QRunnable):
    def __init__(self, file_path, overwrite, date, seq_no_option, custom_prefix):
        super().__init__()
        self.file_path = file_path
        self.overwrite = overwrite
        self.date = date
        self.seq_no_option = seq_no_option
        self.custom_prefix = custom_prefix
        self.signals = WorkerSignals()

    def run(self):
        try:
            self.signals.progress.emit(0)  # 更新進度
            request = FileMetadata(self.file_path, self.date, self.seq_no_option, self.custom_prefix)
            cleaned_data = DataManagerService.load_and_prepare_data(request)
            self.signals.progress.emit(20)  # 更新進度

            if cleaned_data is None:
                raise ValueError("Failed to process data")

            doc = DocumentManagerService.create_document()
            column_widths = (
                Inches(0.8), Inches(0.5), Inches(1.962), Inches(1.24), Inches(0.2), Inches(0.3), Inches(0.2))
            TableFormatterService.create_and_format_table(doc, cleaned_data, column_widths)
            self.signals.progress.emit(70)  # 更新進度

            output_path = self.file_path.replace('.xlsx', '.docx')
            if not self.overwrite:
                output_path = event_handlers.handle_existing_file(output_path)

            DocumentManagerService.save_document(doc, output_path)
            self.signals.progress.emit(100)  # 更新進度
            self.signals.finished.emit(self.file_path, output_path)
        except Exception as e:
            self.signals.error.emit(self.file_path, str(e))
