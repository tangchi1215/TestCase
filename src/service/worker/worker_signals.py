from PyQt6.QtCore import QObject, pyqtSignal


class WorkerSignals(QObject):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str, str)
    error = pyqtSignal(str, str)
