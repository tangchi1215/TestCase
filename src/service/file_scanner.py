import glob
import os


class FileScanner:
    @staticmethod
    def scan_xlsx_files(directory):
        path_pattern = os.path.join(directory, '*.xlsx')
        return glob.glob(path_pattern)
