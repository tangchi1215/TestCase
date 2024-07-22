import sys
from PyQt6 import QtWidgets, QtGui
from PyQt6.QtCore import QTimer
from ui.slash_screen import SplashScreen
from ui.draggable_window import DraggableWindow
from utils.resource_path import resource_path

# SPLASH_IMAGE_PATH = "D:\\TestCase\\src\\assets\\img\\splash_image.png"

SPLASH_IMAGE_PATH = "./src/assets/img/splash_image.png"


def show_main_window():
    window = DraggableWindow()
    window.show()


def show_splash_screen():
    app = QtWidgets.QApplication(sys.argv)
    splash_pix = QtGui.QPixmap(resource_path(SPLASH_IMAGE_PATH))
    splash = SplashScreen(splash_pix)
    splash.show()

    # 模擬初始化時間
    QTimer.singleShot(3000, splash.close)  # 啟動畫面顯示 3 秒
    QTimer.singleShot(3000, show_main_window)  # 啟動畫面關閉後顯示主窗口

    sys.exit(app.exec())


if __name__ == "__main__":
    show_splash_screen()
