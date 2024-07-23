import sys

from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import QApplication
from ui.slash_screen import SplashScreen
from utils.resource_path import resource_path

# SPLASH_IMAGE_PATH = "D:\\TestCase\\src\\assets\\img\\splash_image.png"
SPLASH_IMAGE_PATH = "./src/assets/img/splash_image.png"


def perform_initialization():
    import time
    time.sleep(3)


class InitializationThread(QThread):
    initialized = pyqtSignal()

    def run(self):
        perform_initialization()
        self.initialized.emit()


def show_main_window():
    from ui.draggable_window import DraggableWindow
    window = DraggableWindow()
    window.show()


def show_splash_screen():
    app = QApplication(sys.argv)
    splash_pix = QPixmap(resource_path(SPLASH_IMAGE_PATH))
    splash = SplashScreen(splash_pix)
    splash.show()

    init_thread = InitializationThread()
    init_thread.initialized.connect(lambda: (splash.close(), show_main_window()))
    init_thread.start()

    sys.exit(app.exec())


if __name__ == "__main__":
    show_splash_screen()
