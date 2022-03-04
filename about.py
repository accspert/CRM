from PyQt5 import uic
from PyQt5.QtWidgets import QWidget
from PyQt5 import QtCore

class Ui_About(QWidget):
    def __init__(self):
        super(Ui_About, self).__init__()
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        uic.loadUi(r"about.ui", self)
