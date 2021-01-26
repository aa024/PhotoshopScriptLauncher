# -*- coding: utf-8 -*-
import sys
import os
from PySide2.QtWidgets import *
from PySide2.QtGui import *
from PySide2.QtCore import *
import win32com.client


class MainWindow(QMainWindow):
    def __init__(self, app):
        super(MainWindow, self).__init__()
        self.app = app
        self.setWindowTitle("Photoshop Script Launcher")
        self.path = sys.argv[1]

    def main(self):
        scroll_area = QScrollArea()
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignTop)
        group_box = QGroupBox(os.path.basename(self.path))
        group_box.setLayout(layout)
        self.create_layout(self.path, layout)
        scroll_area.setWidget(group_box)
        self.setCentralWidget(scroll_area)
        self.show()
        self.app.exec_()

    def create_layout(self, path, layout):
        main_layout = QHBoxLayout()
        list_dir = []
        list_file = []
        for e in os.listdir(path):
            path_e = os.path.join(path, e)
            if os.path.isdir(path_e):
                list_dir.append(path_e)
            else:
                list_file.append(path_e)
        for f in list_file:
            if f[-4:] == ".jsx" or f[-3:] == ".js":
                self.add_button(f, layout)

        layout.addLayout(main_layout)
        for d in list_dir:
            layout2 = QVBoxLayout()
            layout2.setAlignment(Qt.AlignTop)
            group_box = QGroupBox(os.path.basename(d))

            self.create_layout(d, layout2)
            group_box.setLayout(layout2)
            main_layout.addWidget(group_box)

    def add_button(self, file_path, layout):
        btn = QPushButton(os.path.basename(file_path))
        btn.setMaximumWidth(512)
        btn.clicked.connect(lambda: self.on_click(file_path))
        layout.addWidget(btn)

    def on_click(self, s):
        ps.DoJavaScriptFile(s)


def main():
    app = QApplication()
    window = MainWindow(app)
    window.main()


if __name__ == "__main__":
    ps = win32com.client.Dispatch("Photoshop.Application")
    main()
