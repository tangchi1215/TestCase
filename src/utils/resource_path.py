import sys
import os

def resource_path(relative_path):
    """ 獲取資源的絕對路徑，適用於開發環境和打包後環境 """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)
