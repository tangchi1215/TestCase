import datetime


class FileMetadata:
    def __init__(self, file_path, test_date=None):
        self.file_path = file_path
        self.test_date = test_date if test_date is not None else datetime.datetime.now()
