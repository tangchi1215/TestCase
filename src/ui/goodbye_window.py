from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel
from PyQt6.QtGui import QMovie
from PyQt6.QtMultimedia import QMediaPlayer, QAudioOutput
from PyQt6.QtCore import QUrl, QTimer, pyqtSignal, Qt


class GoodbyeWindow(QDialog):
    finished = pyqtSignal()

    def __init__(self, gif_path, sound_path, parent=None):
        super().__init__(parent)
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        size = self.geometry()
        self.setGeometry((screen.width() - size.width()) // 2, (screen.height() - size.height()) // 2,
                         500, 500)

        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)

        layout = QVBoxLayout()
        self.setLayout(layout)

        label = QLabel(self)
        label.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        layout.addWidget(label)
        layout.addWidget(label)
        movie = QMovie(gif_path)
        label.setMovie(movie)
        movie.start()

        self.player = QMediaPlayer(self)
        audio_output = QAudioOutput(self)
        self.player.setAudioOutput(audio_output)
        self.player.setSource(QUrl.fromLocalFile(sound_path))
        self.player.play()

        self.player.mediaStatusChanged.connect(self.check_media_status)
        movie.finished.connect(self.finish)

    def check_media_status(self, status):
        if status == QMediaPlayer.MediaStatus.EndOfMedia:
            self.finish()

    def finish(self):
        QTimer.singleShot(1000, self.close)
        self.finished.emit()
