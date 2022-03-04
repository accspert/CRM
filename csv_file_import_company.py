from PyQt5 import QtCore, QtGui, QtWidgets
import pandas as pd
from PyQt5.QtWidgets import QWidget
from PyQt5.uic import loadUiType
from PyQt5 import QtCore
from PyQt5.QtWidgets import *

from PyQt5.QtWidgets import QHeaderView, QMessageBox, QProgressDialog

import PandasTableModel
import sys
from sql import SqlHelper

Ui_import, _ = loadUiType('import_company.ui')

class Ui_csv_file_import_company(QMainWindow, Ui_import): 
    def __init__(self, helper):
        self.helper = helper
        QWidget.__init__(self)
        self.setupUi(self) 
        self.data = None
        self.df = None        

        self.pushButtonLoadCSV.clicked.connect(self.selectFile)
        self.pushButtonSaveCSV.clicked.connect(self.saveTodb)
        self.pushButtonClose.clicked.connect(self.close)

    def saveTodb(self):
        if self.df is None:
            error_dialog = QtWidgets.QErrorMessage()
            error_dialog.showMessage('No CSV File Selected')
            error_dialog.exec()
            
        try:
            textboxes_company = [self.lineEditCompanyName, self.lineEditCompanyAddress, self.lineEditCompanyPhone, 
                                 self.lineEditCompanyWebsite,self.lineEditCompanyEmail]
    
            txtvalues_company = [x.text() for x in textboxes_company if x.text()]
    
            if hasDuplicate(txtvalues_company):
                error_dialog = QtWidgets.QErrorMessage()
                error_dialog.showMessage('Entries Contain Duplicate Column Numbers')
                error_dialog.exec()
            elif hasNonDigit(txtvalues_company):
                error_dialog = QtWidgets.QErrorMessage()
                error_dialog.showMessage('Only Numbers are Allowed')
                error_dialog.exec()
            else:
                
                txtvalues_company = [x.text() for x in textboxes_company]
    
                for row in self.data.iterrows():
                    db_row_company =[]
                    for i in txtvalues_company:
                        if i =='':
                            db_row_company.append('')
                        else:
                            db_row_company.append(row[1][int(i)])
                            
                    record = tuple(db_row_company)
                    self.helper.insert("INSERT INTO companies (companyname, mergedaddress, phone, website, email) VALUES(?,?,?,?,?)",record)
        except Exception as e:
            QMessageBox.warning(self, 'Error', str(e), QMessageBox.Ok)
            return
        self.statusBar().showMessage('Entries saved',5000)

    def selectFile(self):
        fname = QtWidgets.QFileDialog.getOpenFileName(QtWidgets.QWidget(), 'Open a file', '',
                                                      'CSV Files (*.csv)')
        try:
            self.data = pd.read_csv(fname[0])
            cols = list(self.data)
            self.df = pd.DataFrame(columns=['Columns'])
            for i in cols:
                self.df = self.df.append({'Columns': i}, ignore_index=True)
            model = PandasTableModel.PandasTableModel(self.df)
            self.tableView.setModel(model)
        except Exception as e:
            QMessageBox.warning(self, 'Error', str(e), QMessageBox.Ok)
            return


def hasDuplicate(listOfElems):
    """
      Check if given list contains any duplicates
    """
    if len(listOfElems) == len(set(listOfElems)):
        return False
    else:
        return True


def hasEmpty(listofElems):
    """
      Check if given list contains any empty
    """
    for i in listofElems:
        if i == "":
            return True
    return False


def hasNonDigit(listofElems):
    for x in listofElems:
        if x.isdigit() or x == "":
            return False
    return True