from PyQt5 import uic
from PyQt5.QtWidgets import QWidget
from PyQt5 import QtCore


# Main Window Class
class Ui_available(QWidget):
    def __init__(self):
        super(Ui_available, self).__init__()
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        uic.loadUi(r"available.ui", self)
     