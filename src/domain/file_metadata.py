import datetime


class FileMetadata:
    def __init__(self, file_path, test_date=None, seq_no_option="filename", custom_prefix=""):
        self.file_path = file_path
        self.test_date = test_date if test_date is not None else datetime.datetime.now()
        self.seq_no_option = seq_no_option
        self.custom_prefix = custom_prefix
