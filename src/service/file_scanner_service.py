import glob
import os


class FileScannerService:
    @staticmethod
    def scan_xlsx_files(directory):
        path_pattern = os.path.join(directory, '*.xlsx')
        return glob.glob(path_pattern)
