import os
import sys
import time

from PyQt5.QtCore import QDir
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QFileSystemModel, QComboBox, QTreeView, \
    QWidget, QLabel
from PyQt5 import QtGui
from PyQt5 import QtCore
from win32comext.shell.demos.servers.shell_view import FileSystemView

from src.run_pdf_xlsx import pdf_to_omie_xlsx


class Window(QMainWindow):
    def __init__(self, dir_path):
        super().__init__()

        self.top = 600
        self.left = 200
        self.large = 600
        self.high = 300
        self.title = 'Janela'
        # centralWidget = QWidget()
        # self.setCentralWidget(centralWidget)

        self.model = QFileSystemModel()
        self.model.setRootPath(dir_path)

        self.path = ''

        self.tree = QTreeView()
        self.tree.setModel(self.model)
        self.tree.clicked.connect(self.onClicked)
        self.tree.setRootIndex(self.model.index(dir_path))
        self.tree.setAnimated(False)
        self.tree.setIndentation(20)
        self.tree.setSortingEnabled(True)
        # self.tree.resize(100,100)
        # self.tree.move(100, 100)
        self.buttonload = QPushButton('Carregar', self)
        self.buttonload.move(80, 50)
        self.buttonload.resize(80, 50)
        self.buttonload.setStyleSheet('QPushButton {background-color:#0FB328;font:bold}')
        self.buttonload.clicked.connect(self.buttonload_click)

        self.label_1 = QLabel(self)
        self.label_1.setText('Selecione um arquivo .PDF')
        self.label_1.move(50, 25)
        self.label_1.resize(300, 25)
        self.label_1.setStyleSheet('QLabel {font:bold;font-size:15px;color:"blue"}')

        # self.layout = QVBoxLayout(centralWidget)
        # self.layout.addWidget(self.buttonload)
        # self.layout.addWidget(self.tree)
        # self.setLayout(self.layout)

        self.setup()
        self.load_window()

    def buttonload_click(self):
        if len(self.path) < 4:
            self.label_1.setText('Selecione um arquivo PDF valido!')
        elif self.path[-4:] != '.pdf' and self.path[-4:] != '.PDF':
            self.label_1.setText('Selecione um arquivo PDF valido!')
        else:
            self.label_1.setText("Aguarde...")
            self.label_1.setText(pdf_to_omie_xlsx(self.path, os.getcwd() + "\Omie.xlsx", self.label_1))
            # time.sleep(10)
            # self.label_1.setText('Selecione um arquivo .PDF')

    def load_window(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.large, self.high)
        self.show()

    def setup(self):
        centralWidget = QWidget()
        self.setCentralWidget(centralWidget)
        layout = QVBoxLayout(centralWidget)
        layout.addWidget(self.buttonload)
        layout.addWidget(self.label_1)
        layout.addWidget(self.tree)
        self.setLayout(layout)

    def onClicked(self, index):
        self.label_1.setText('Selecione um arquivo .PDF')
        self.path = self.sender().model().filePath(index)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    dir_path = os.getcwd()
    # j = Window()
    demo = Window(dir_path)
    demo.show
    sys.exit(app.exec())
