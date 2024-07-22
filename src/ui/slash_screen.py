from PyQt6 import QtWidgets, QtCore


class SplashScreen(QtWidgets.QSplashScreen):
    def __init__(self, pixmap):
        scaled_pixmap = pixmap.scaled(800, 500, QtCore.Qt.AspectRatioMode.KeepAspectRatio,
                                      QtCore.Qt.TransformationMode.SmoothTransformation)
        super().__init__(scaled_pixmap, QtCore.Qt.WindowType.WindowStaysOnTopHint)
        self.setMask(scaled_pixmap.mask())
