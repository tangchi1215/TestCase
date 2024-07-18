def load_style(widget, style_path):
    """ 加載 qss 樣式並應用到 widget """
    with open(style_path, "r") as f:
        widget.setStyleSheet(f.read())
