from PyQt5.QtWidgets import QDialog, QDialogButtonBox, QVBoxLayout, QLabel, QTableWidgetItem
from PyQt5 import QtGui


class CustomDialog(QDialog):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("About CRM")

        QBtn = QDialogButtonBox.Ok | QDialogButtonBox.Cancel

        self.buttonBox = QDialogButtonBox(QBtn)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)

        self.layout = QVBoxLayout()
        message = QLabel("Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has "
                         "\nbeen the industry's standard dummy text ever since the 1500s, "
                         "\nwhen an unknown printer took a galley of type and scrambled it to make a type specimen "
                         "\nbook. It has survived not only five centuries, "
                         "\nbut also the leap into electronic typesetting, remaining essentially unchanged. It was "
                         "\npopularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum "
                         "\npassages, and more recently with desktop publishing software like Aldus PageMaker including "
                         "\nversions of Lorem Ipsum.")
        self.layout.addWidget(message)
        self.layout.addWidget(self.buttonBox)
        self.setLayout(self.layout)


class TheIconItem(QTableWidgetItem):
    def __init__(self, icon_name, icon):
        super(TheIconItem, self).__init__()
        self.icon_name = icon_name
        self.setIcon(QtGui.QIcon(icon))