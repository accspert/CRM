import sys
from os.path import exists
from os import mkdir
import os
from datetime import time, date
from datetime import datetime
import ctypes
import csv

from PyQt5 import uic, QtGui, QtCore
from PyQt5.QtWidgets import *

from custom import CustomDialog, TheIconItem
from folder_selector import Ui_FolderSelector
from about import Ui_About
from available import Ui_available
from csv_file_import_contact import Ui_csv_file_import_contact
from csv_file_import_company import Ui_csv_file_import_company

from sql import SqlHelper
from ErrorLogger import *
import traceback

from keyvalidator import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

from crypto import *

from dateutil.parser import parse
from operator import itemgetter


# Main Window Class
class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        global helper
        
        try:

            self.folder_selector = None
            self.about_dialog = None
            self.available_dialog = None
            
            uic.loadUi(r"MainWindow.ui", self)
            
            current_db_file = open(r"assets/db/current_db.txt", "r+")
            db_path = current_db_file.readline()
            current_db_file.close()
            
            if os.path.isfile(db_path): 
                helper = SqlHelper(db_path)
            else:
                self.open_database('')
            
            self.set_date_time()
            self.hide_product_box()
            self.hide_call_box()
            self.hide_event_box()
            self.hide_task_box()
            self.hide_deal_box()
            self.handle_tab_widget()
            self.handle_contact_menu_buttons_do()
            self.handle_main_buttons_do()
            self.refresh()
        except Exception as e:
            ErrorLogger.WriteError(traceback.format_exc())
            QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))          
        

    def refresh(self):
        try:
            self.fill_deal_stage_amount()
            self.fill_stages_label()
            self.fill_cmbBox_adress_type()
            self.fill_cmbBox_group()
            self.fill_cmbBox_status()
            self.fill_cmbBox_products()
            self.fill_cmbBox_products_category()        
            self.fill_cmbBox_groups()
            self.fill_cmbBox_companies()
            self.fill_cmbBox_deals()
            self.fill_cmbBox_contacts()
            self.fill_cmbBox_stages()
            self.fill_cmbBox_activity_filter()
            self.fill_deals_qualification_table()
            self.fill_deals_needs_analyse_table()
            self.fill_deals_offer_table()
            self.fill_deals_negotiation_table()
            self.fill_deals_closed_table()
            self.fill_groups_table() 
            self.fill_status_table() 
            self.fill_contact_table()
            self.fill_company_table()
            self.fill_products_table()
            self.fill_product_category_table()         
            self.fill_tasks_table()
            self.fill_events_table()
            self.fill_calls_table()
            self.clear_product_category_fields()
            self.clear_status_fields()
            self.clear_company_address_fields()
            self.clear_product_fields()
            self.clear_company_fields()
            self.clear_contact_available_address_table()
            self.clear_contact_address_fields()
            self.clear_call_fields()
            self.clear_event_fields()
            self.clear_deal_fields()
            self.clear_contact_fields()
            self.clear_task_fields()
            self.clear_group_fields()        
        except Exception as e:
            ErrorLogger.WriteError(traceback.format_exc())
            QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))          

    def handle_tab_widget(self):
        self.tabWidget.tabBar().setVisible(False)
        self.tabWidget.setCurrentIndex(0)

    def handle_contact_menu_buttons_do(self):
        self.actionNew_Database.triggered.connect(self.new_database)
        self.actionOpen_Database.triggered.connect(self.open_database)
        self.actionAbout_CRM.triggered.connect(self.about_CRM)

    def handle_main_buttons_do(self):
        self.label_22.linkActivated.connect(lambda: self.go_to_contact(self.comboBox_10)) 
        self.label_37.linkActivated.connect(lambda: self.go_to_contact(self.comboBox_19))
        self.label_45.linkActivated.connect(lambda: self.go_to_contact(self.comboBox_22))
        self.label_76.linkActivated.connect(lambda: self.go_to_contact(self.comboBox_9))
        
        self.label_98.linkActivated.connect(lambda: self.go_to_company(self.comboBox_17))
        self.label_100.linkActivated.connect(lambda: self.go_to_company(self.comboBox_21))
        self.label_102.linkActivated.connect(lambda: self.go_to_company(self.comboBox_24))
        self.label_6.linkActivated.connect(lambda: self.go_to_company(self.comboBox_3))
        self.label_77.linkActivated.connect(lambda: self.go_to_company(self.comboBox_7))
        
        self.label_97.linkActivated.connect(lambda: self.go_to_deal(2,self.comboBox_12))
        self.label_99.linkActivated.connect(lambda: self.go_to_deal(2,self.comboBox_20))
        self.label_101.linkActivated.connect(lambda: self.go_to_deal(2,self.comboBox_23))
        
        self.comboBox_2.view().pressed.connect(self.filter_by_group)
        self.comboBox_25.view().pressed.connect(self.filter_event)
        self.comboBox_26.view().pressed.connect(self.filter_task)
        self.comboBox_27.view().pressed.connect(self.filter_call)
        self.comboBox.view().pressed.connect(self.filter_by_status)
        self.comboBox_4.currentIndexChanged.connect(self.sort_contact_table)
        
        self.lineEdit.installEventFilter(self)
        self.lineEdit.textEdited.connect(self.filter_by_text)
        
        self.pushButton.clicked.connect(self.delete_contact)
        self.pushButton_2.clicked.connect(self.show_contacts)
        self.pushButton_3.clicked.connect(self.filter_no_filter)
        self.pushButton_4.clicked.connect(self.show_activity)
        self.pushButton_5.clicked.connect(self.export_as_csv)
        self.pushButton_6.clicked.connect(self.show_admin)
        self.pushButton_7.clicked.connect(self.save_contact)
        self.pushButton_8.clicked.connect(lambda: self.go_to_activity(16))
        self.pushButton_9.clicked.connect(self.show_csv_window_contact)
        self.pushButton_10.clicked.connect(self.show_products)
        self.pushButton_11.clicked.connect(self.save_task)
        self.pushButton_12.clicked.connect(self.delete_task)
        self.pushButton_13.clicked.connect(self.save_event)
        self.pushButton_14.clicked.connect(self.save_group)
        self.pushButton_15.clicked.connect(self.show_deals)
        self.pushButton_16.clicked.connect(self.clear_group_fields)
        self.pushButton_17.clicked.connect(self.delete_group)
        self.pushButton_18.clicked.connect(self.hide_task_box)
        self.pushButton_19.clicked.connect(self.clear_contact_address_fields)
        self.pushButton_20.clicked.connect(self.show_companies)
        self.pushButton_21.clicked.connect(self.save_status)
        self.pushButton_22.clicked.connect(self.clear_status_fields)
        self.pushButton_23.clicked.connect(self.delete_status)
        self.pushButton_24.clicked.connect(self.show_contact_detail)
        self.pushButton_25.clicked.connect(self.clear_task_fields)
        self.pushButton_26.clicked.connect(self.show_task_box)
        self.pushButton_27.clicked.connect(self.show_event_box)
        self.pushButton_28.clicked.connect(self.clear_event_fields)
        self.pushButton_29.clicked.connect(self.hide_event_box)
        self.pushButton_30.clicked.connect(self.delete_event)
        self.pushButton_31.clicked.connect(self.save_call)
        self.pushButton_32.clicked.connect(self.clear_call_fields)
        self.pushButton_33.clicked.connect(self.hide_call_box)
        self.pushButton_34.clicked.connect(self.delete_call)
        self.pushButton_35.clicked.connect(self.show_call_box)
        self.pushButton_36.clicked.connect(self.show_company_detail)
        self.pushButton_37.clicked.connect(self.delete_company)
        self.pushButton_38.clicked.connect(self.save_company)
        self.pushButton_39.clicked.connect(self.clear_contact_fields)
        self.pushButton_39.clicked.connect(self.clear_contact_address_fields)
        self.pushButton_39.clicked.connect(self.clear_contact_available_address_table)
        self.pushButton_39.clicked.connect(self.clear_contact_activity_table)
        self.pushButton_40.clicked.connect(self.save_stages)
        self.pushButton_41.clicked.connect(self.show_product_box)
        self.pushButton_42.clicked.connect(self.delete_product)
        self.pushButton_43.clicked.connect(self.save_product)
        self.pushButton_44.clicked.connect(self.clear_product_fields)
        self.pushButton_45.clicked.connect(self.hide_product_box)
        self.pushButton_46.clicked.connect(self.delete_contact_address)
        self.pushButton_47.clicked.connect(self.clear_company_fields)
        self.pushButton_47.clicked.connect(self.clear_company_activity_table)
        self.pushButton_48.clicked.connect(self.close)
        self.pushButton_49.clicked.connect(self.save_deal)
        self.pushButton_50.clicked.connect(self.clear_deal_fields)
        self.pushButton_51.clicked.connect(self.hide_deal_box)
        self.pushButton_52.clicked.connect(self.show_deal_box)
        self.pushButton_53.clicked.connect(self.delete_deal)
        self.pushButton_54.clicked.connect(self.save_product_category)
        self.pushButton_55.clicked.connect(self.clear_product_category_fields)
        self.pushButton_56.clicked.connect(self.delete_product_category)
        self.pushButton_57.clicked.connect(lambda: self.go_to_deal(1, None))
        self.pushButton_58.clicked.connect(self.show_csv_window_company)
        self.pushButton_59.clicked.connect(lambda: self.go_to_activity(17))
        self.pushButton_60.clicked.connect(self.go_to_last_contact)
        self.pushButton_61.clicked.connect(self.go_to_next_contact)
        
        self.tableWidget.clicked.connect(self.move_table_fields_contact)
        self.tableWidget_2.clicked.connect(self.move_table_fields_groups) 
        self.tableWidget_3.clicked.connect(self.move_table_fields_address) #Contact Available Address Table
        self.tableWidget_4.clicked.connect(lambda: [self.get_dealID(4), self.show_deal_box()]) 
        self.tableWidget_6.clicked.connect(lambda: [self.get_dealID(6), self.show_deal_box()]) 
        self.tableWidget_7.clicked.connect(lambda: [self.get_dealID(7), self.show_deal_box()]) 
        self.tableWidget_8.clicked.connect(lambda: [self.get_dealID(8), self.show_deal_box()]) 
        self.tableWidget_9.clicked.connect(lambda: [self.get_dealID(9), self.show_deal_box()]) 
        self.tableWidget_5.clicked.connect(self.move_table_fields_status)
        self.tableWidget_13.clicked.connect(self.move_table_fields_company)
        self.tableWidget_14.clicked.connect(self.move_table_fields_product) 
        self.tableWidget_10.clicked.connect(self.move_table_fields_task)
        self.tableWidget_11.clicked.connect(self.move_table_fields_event)
        self.tableWidget_12.clicked.connect(self.move_table_fields_call)
        self.tableWidget_15.clicked.connect(self.move_table_fields_product_category)

  ################################################# @show @hide
        
    def show_deal_box(self):
        self.groupBox_11.show()  
    def hide_deal_box(self):
        self.groupBox_11.hide()
    def show_product_box(self):
        self.groupBox_10.show()  
    def hide_product_box(self):
        self.groupBox_10.hide()  
    def show_company_detail(self):
        self.tabWidget_5.setCurrentIndex(1)  
    def show_call_box(self):
        self.groupBox_7.show()  
    def hide_call_box(self):
        self.groupBox_7.hide()
        self.clear_call_fields()
    def show_event_box(self):
        self.groupBox_6.show()  
    def hide_event_box(self):
        self.groupBox_6.hide()
        self.clear_event_fields()
    def show_contact_detail(self):
        self.tabWidget_4.setCurrentIndex(1)
    def show_task_box(self):
        self.groupBox_4.show()
    def hide_task_box(self):
        self.groupBox_4.hide()
        self.clear_task_fields()
    def show_deals(self):
        self.tabWidget.setCurrentIndex(0)  
    def show_contacts(self):
        self.tabWidget.setCurrentIndex(1)
        # self.tabWidget_4.setCurrentIndex(0)
    def show_companies(self):
        self.tabWidget.setCurrentIndex(2)
        self.tabWidget_5.setCurrentIndex(0)
    def show_products(self):
        self.tabWidget.setCurrentIndex(3)
    def show_activity(self):
        self.tabWidget.setCurrentIndex(4)        
    def show_admin(self):
        self.tabWidget.setCurrentIndex(5)
################################################### @sort
    def sort_contact_table(self, index):
        self.tableWidget.sortItems(index -1)
  ################################################# @go_to
    def go_to_last_contact(self):
        if self.lineEdit_11.text():
            row = self.tableWidget.currentRow() -1
            self.tableWidget.selectRow(row)
            self.move_table_fields_contact()
    def go_to_next_contact(self):
        if self.lineEdit_11.text():
            row = self.tableWidget.currentRow() +1
            self.tableWidget.selectRow(row)
            self.move_table_fields_contact()
    def go_to_company(self, cmbBox):
        if cmbBox.currentIndex() == 0:
            return
        status_name = str(cmbBox.currentText())
        self.tabWidget.setCurrentIndex(2)
        self.show_company_detail()
        data = helper.select(f"select * from Companies where CompanyName = '{status_name}'")
        if data:
            try:
                self.clear_company_fields()
                self.clear_company_address_fields()
                self.lineEdit_25.setText(str(data[0][0]))
                self.lineEdit_26.setText(data[0][1])
                self.lineEdit_27.setText(data[0][2])
                self.lineEdit_28.setText(data[0][3])
                self.plainTextEdit_8.setPlainText(data[0][4])
                self.fill_company_activity_table()

            except Exception as e:
                        ErrorLogger.WriteError(traceback.format_exc())
                        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))    
    def go_to_contact(self, cmbBox):
        if cmbBox.currentIndex() == 0:
            return
        status_name = str(cmbBox.currentText())
        self.tabWidget.setCurrentIndex(1)
        self.show_contact_detail()
        data = helper.select(f"select * from ContactQ2 where CS = '{status_name}'")
        if data:
            try:
                self.clear_contact_fields()
                self.clear_contact_address_fields()
                
                self.lineEdit_11.setText(str(data[0][0]))
                self.lineEdit_4.setText(data[0][1])
                self.lineEdit_2.setText(data[0][2])
                self.lineEdit_39.setText(data[0][3])
                self.lineEdit_3.setText(data[0][4])
                self.lineEdit_5.setText(data[0][5])
                self.lineEdit_6.setText(data[0][6])
                self.comboBox_13.setCurrentText(data[0][7])
                self.comboBox_3.setCurrentText(data[0][8])
                self.comboBox_14.setCurrentText(data[0][9])
                self.comboBox_5.setCurrentText(data[0][10])
                self.plainTextEdit.setPlainText(data[0][11])
                
                self.fill_table_available_addresses()
                self.fill_contact_activity_table()
            except Exception as e:
                        ErrorLogger.WriteError(traceback.format_exc())
                        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))    
        
    def go_to_deal(self, caller, cmbBox):
        try:
            if caller == 1 and len(self.tableWidget.selectionModel().selectedRows()) >0:
                contactId = self.tableWidget.item(self.tableWidget.currentRow(),0).text()
                data = helper.select(f"select dealId from deals where contactId = {contactId}")
            
            if caller == 2:
                deal_name_cmb = cmbBox.currentText()
                data= helper.select(f"select dealid from deals where dealname = '{deal_name_cmb}'")
    
            if data:
                dealId = data[0][0]
                self.show_deals()
                self.show_deal_box()            
                self.move_table_fields_deal(dealId)
            else:
                QtWidgets.QMessageBox.information(None, 'No Data', 'No Deal with this contact') 
        except Exception as e:
                ErrorLogger.WriteError(traceback.format_exc())
                QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e)) 

    def go_to_activity(self, caller):
        
        if caller == 17:
            if len(self.tableWidget_17.selectionModel().selectedRows()) == 0:
                QtWidgets.QMessageBox.information(None, 'No Selection', 'Select one row')
                return
            else: 
                activity_type = self.tableWidget_17.item(self.tableWidget_17.currentRow(),0).text()
                activityId   = int(self.tableWidget_17.item(self.tableWidget_17.currentRow(),1).text())
        elif caller == 16:
            if len(self.tableWidget_16.selectionModel().selectedRows()) ==0:
                QtWidgets.QMessageBox.information(None, 'No Selection', 'Select one row')
                return
            else: 
                activity_type = self.tableWidget_16.item(self.tableWidget_16.currentRow(),0).text()
                activityId   = int(self.tableWidget_16.item(self.tableWidget_16.currentRow(),1).text())
            
        try:
            if activity_type == 'Task':
                task_data = helper.select(f"""select TaskID, TaskName, DueDate, StatusName, DealName,
                                          CompanyName, HighPriority,Completed,Description
                                          from TaskQ where taskId = {activityId}""")
                self.show_activity()
                self.tabWidget_2.setCurrentIndex(0)
                self.show_task_box()
                self.clear_task_fields()
                self.lineEdit_8.setText(str(task_data[0][0]))
                self.lineEdit_7.setText(str(task_data[0][1]))
                self.dateEdit.setDate(parse(task_data[0][2], dayfirst=True))
                self.comboBox_10.setCurrentText(task_data[0][3])
                self.comboBox_12.setCurrentText(task_data[0][4])
                self.comboBox_17.setCurrentText(task_data[0][5])
                self.checkBox.setChecked(task_data[0][6]) 
                self.checkBox_2.setChecked(task_data[0][7])
                self.plainTextEdit_2.setPlainText(task_data[0][8])
            
            if activity_type == 'Event':
                event_data = helper.select(f"""select EventID, EventTitle, DateFrom, TimeFrom, DateTo, TimeTo, Location,
                                          StatusName, DealName, CompanyName, Participants, Description
                                          from EventQ where EventId = {activityId}""")
                self.show_activity()
                self.tabWidget_2.setCurrentIndex(1)
                self.show_event_box()
                self.clear_event_fields()
                
                self.lineEdit_17.setText(str(event_data[0][0]))
                self.lineEdit_21.setText(str(event_data[0][1]))
                self.dateEdit_6.setDate(parse(event_data[0][2], dayfirst=True))
                
                dateD = parse(event_data[0][3])
                dateK = dateD.time()
                self.timeEdit_2.setTime(dateK)
                
                self.dateEdit_2.setDate(parse(event_data[0][4], dayfirst=True))
                
                dateD = parse(event_data[0][5])
                dateK = dateD.time()
                self.timeEdit_3.setTime(dateK)
                
                self.lineEdit_38.setText(str(event_data[0][6]))                    
                self.comboBox_19.setCurrentText(event_data[0][7])
                self.comboBox_20.setCurrentText(event_data[0][8])
                self.comboBox_21.setCurrentText(event_data[0][9])
                self.lineEdit_42.setText(str(event_data[0][10]))
                self.plainTextEdit_3.setPlainText(event_data[0][11])
                
            if activity_type == 'Call':
                call_data = helper.select(f"""select CallID, ToFrom, CallTime, CallDate, CallType,
                                          StatusName, DealName, CompanyName, CallAgenda
                                          from CallQ where CallId = {activityId}""")
                self.show_activity()
                self.tabWidget_2.setCurrentIndex(2)
                self.show_call_box()
                self.clear_call_fields()
                
                self.lineEdit_22.setText(str(call_data[0][0]))
                self.lineEdit_46.setText(str(call_data[0][1]))
                dateD = parse(call_data[0][2])
                dateK = dateD.time()
                self.timeEdit.setTime(dateK)
                
                self.dateEdit_4.setDate(parse(call_data[0][3], dayfirst=True))
                self.comboBox_11.setCurrentText(call_data[0][4])
                self.comboBox_22.setCurrentText(call_data[0][5])
                self.comboBox_23.setCurrentText(call_data[0][6])
                self.comboBox_24.setCurrentText(call_data[0][7])
                self.plainTextEdit_4.setPlainText(call_data[0][8])
        except Exception as e:
                ErrorLogger.WriteError(traceback.format_exc())
                QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e)) 

############################################################ @fill        
    def fill_deal_stage_amount(self):
        deal_stage_1_amount = str(helper.select("select sum(amount) from deals where stageid = 1")[0][0])
        self.label_25.setText("Amount: " + deal_stage_1_amount)
        
        deal_stage_2_amount = str(helper.select("select sum(amount) from deals where stageid = 2")[0][0])
        self.label_43.setText("Amount: " + deal_stage_2_amount)
        
        deal_stage_3_amount = str(helper.select("select sum(amount) from deals where stageid = 3")[0][0])
        self.label_47.setText("Amount: " + deal_stage_3_amount)
        
        deal_stage_4_amount = str(helper.select("select sum(amount) from deals where stageid = 4")[0][0])
        self.label_112.setText("Amount: " + deal_stage_4_amount)
        
        deal_stage_5_amount = str(helper.select("select sum(amount) from deals where stageid = 5")[0][0])
        self.label_113.setText("Amount: " + deal_stage_5_amount)
    
    def fill_table_widget(self, table_widget, sql_statment, sort_column):
        try:
            table_widget.setRowCount(0)        
            data = helper.select(sql_statment)
            for row, form in enumerate(data):
                row_position = table_widget.rowCount()
                table_widget.insertRow(row_position)  
                for column, item in enumerate(form):
                    if item:
                        table_widget.setItem(row, column, QTableWidgetItem(str(item)))
            table_widget.sortItems(sort_column, order=0) #Ascending
                    
        except Exception as e:
                ErrorLogger.WriteError(traceback.format_exc())
                QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e)) 
         
    def fill_contact_activity_table(self):
        try:
            if self.lineEdit_11.text():
                all_data = []
                contactId = self.lineEdit_11.text()
                task_data = helper.select(f"select taskid, duedate, taskname from tasks where relatedtocontact={contactId}")
                all_data =[('Task',) + x for x in task_data]
                event_data = helper.select("select eventid, datefrom, eventtitle from events where relatedtocontact="+contactId)
                all_data += [('Event',) + x for x in event_data]
                call_data = helper.select("select callid, calldate, tofrom from calls where relatedtocontact="+contactId)
                all_data += [('Call',) + x for x in call_data]
                all_data.sort(key=itemgetter(2))
                self.tableWidget_17.setRowCount(0)
                for row, form in enumerate(all_data):
                    row_position = self.tableWidget_17.rowCount()
                    self.tableWidget_17.insertRow(row_position)
                    for column, item in enumerate(form):
                        if item:
                            self.tableWidget_17.setItem(row, column , QTableWidgetItem(str(item)))
            self.tableWidget_17.setColumnWidth(0, 7);
            self.tableWidget_17.setColumnWidth(1, 3);
        except Exception as e:
                ErrorLogger.WriteError(traceback.format_exc())
                QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e)) 
   
    def fill_company_activity_table(self):
        try:
            if self.lineEdit_25.text():
                all_data = []
                companyId = self.lineEdit_25.text()
                task_data = helper.select("select taskid, duedate, taskname from tasks where relatedtocompany="+companyId)
                all_data =[('Task',) + x for x in task_data]
                event_data = helper.select("select eventid, datefrom, eventtitle from events where relatedtocompany="+companyId)
                all_data += [('Event',) + x for x in event_data]
                call_data = helper.select("select callid, calldate, tofrom from calls where relatedtocompany="+companyId)
                all_data += [('Call',) + x for x in call_data]
                all_data.sort(key=itemgetter(2))
                self.tableWidget_16.setRowCount(0)
                for row, form in enumerate(all_data):
                    row_position = self.tableWidget_16.rowCount()
                    self.tableWidget_16.insertRow(row_position)
                    for column, item in enumerate(form):
                        if item:
                            self.tableWidget_16.setItem(row, column , QTableWidgetItem(str(item)))
            self.tableWidget_16.setColumnWidth(1, 10);
        except Exception as e:
                ErrorLogger.WriteError(traceback.format_exc())
                QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e)) 
        
        
    def fill_contact_table(self):
        self.fill_table_widget(self.tableWidget, "select * from ContactQ",1)
    def fill_company_table(self):
        self.fill_table_widget(self.tableWidget_13, "select * from companyAddresses",1 )
    def fill_tasks_table(self):
        try:
            self.tableWidget_10.setRowCount(0)
            data = helper.select("select * from TaskQ")
            cb = QCheckBox()
            for row, form in enumerate(data):
                row_position = self.tableWidget_10.rowCount()
                self.tableWidget_10.insertRow(row_position)
                for column, item in enumerate(form):
                    if column in [6, 7]:
                        status_icon = None
                        
                        if item :
                            icon = r"resources/icon/check.ico"
                            status_icon = TheIconItem("check", icon)
                        else:
                            icon = r"resources/icon/uncheck.ico"
                            status_icon = TheIconItem("uncheck", icon)
                        
                        self.tableWidget_10.setItem(row, column, status_icon)
                    else:
                        if item:
                            self.tableWidget_10.setItem(row, column, QTableWidgetItem(str(item)))
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))    
    def fill_events_table(self):
        self.fill_table_widget(self.tableWidget_11, "select * from EventQ",2)

    def fill_calls_table(self):
        self.fill_table_widget(self.tableWidget_12, "select * from CallQ",3)
    
    def fill_table_available_addresses(self): 
        try:
            self.tableWidget_3.setRowCount(0)
            contactOn = self.lineEdit_11.text() 
            data = helper.select(f"select AddressID, AddressTypeName, MailingStreet, MailingCity, MailingState, MailingCountry, MailingZip FROM ContactAddresses where ContactID = {contactOn}")
            for row, form in enumerate(data):
                row_position = self.tableWidget_3.rowCount()
                self.tableWidget_3.insertRow(row_position) 
                for column, item in enumerate(form):
                    if item:
                        self.tableWidget_3.setItem(row, column, QTableWidgetItem(str(item)))
            #address fields
            if data:
                result = data.pop()             
                self.lineEdit_40.setText(str((result[0])))
                if result[1]:
                    self.comboBox_15.setCurrentText(str(result[1]))
                self.lineEdit_16.setText(str((result[2])))
                self.lineEdit_15.setText(str((result[3])))
                self.lineEdit_18.setText(str((result[4])))
                self.lineEdit_14.setText(str((result[5])))
                self.lineEdit_20.setText(str((result[6])))
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))    
                
    def fill_products_table(self): 
        try:
            self.tableWidget_14.setRowCount(0)        
            data = helper.select("select * from ProductQ")
            for row, form in enumerate(data):
                row_position = self.tableWidget_14.rowCount()
                self.tableWidget_14.insertRow(row_position)               
                for column, item in enumerate(form):
                    if column in [6]:
                        status_icon = None
    
                        if item:
                            icon = r"resources/icon/check.ico"
                            status_icon = TheIconItem("check", icon)
                        else:
                            icon = r"resources/icon/uncheck.ico"
                            status_icon = TheIconItem("uncheck", icon)                    
                        self.tableWidget_14.setItem(row, column, status_icon)
                    else:
                        if item:
                            self.tableWidget_14.setItem(row, column, QTableWidgetItem(str(item)))
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))    
            
    def fill_deals_qualification_table(self):
        self.fill_table_widget(self.tableWidget_4, "select DealID, DealName from dealsQualificationQ",1 )
        self.tableWidget_4.setColumnWidth(0, 60);
        
    def fill_deals_needs_analyse_table(self):
        self.fill_table_widget(self.tableWidget_7, "select DealID, DealName from dealsNeedsAnalyseQ",1 )
        self.tableWidget_7.setColumnWidth(0, 60);
    def fill_deals_offer_table(self):
        self.fill_table_widget(self.tableWidget_6, "select DealID, DealName from dealsOfferQ",1 )
        self.tableWidget_6.setColumnWidth(0, 60);
    def fill_deals_negotiation_table(self):
        self.fill_table_widget(self.tableWidget_8, "select DealID, DealName from dealsnegotiationQ",1 )
        self.tableWidget_8.setColumnWidth(0, 60);
    def fill_deals_closed_table(self):
        self.fill_table_widget(self.tableWidget_9, "select DealID, DealName from dealsClosedQ",1 )
        self.tableWidget_9.setColumnWidth(0, 60);
    def fill_groups_table(self):
        self.fill_table_widget(self.tableWidget_2, "select * from GroupsQ",1 )
  
    def fill_status_table(self):
        self.fill_table_widget(self.tableWidget_5, "select * from EntityStatus" ,1)
  
    def fill_product_category_table(self):
        self.fill_table_widget(self.tableWidget_15, "select * from productCategory",1 )
    
    def fill_stages_label(self):
        data = helper.select("select StageName from stages")
        
        for i, text_stages in enumerate(data):
            if i ==0:
                self.label_80.setText(text_stages[0])
                self.lineEdit_9.setText(text_stages[0])
            elif i ==1:
                self.label_84.setText(text_stages[0])                
                self.lineEdit_10.setText(text_stages[0])
            elif i ==2:
                self.label_83.setText(text_stages[0])                
                self.lineEdit_43.setText(text_stages[0])
            elif i ==3:
                self.label_82.setText(text_stages[0])                
                self.lineEdit_44.setText(text_stages[0])
            elif i ==4:
                self.label_81.setText(text_stages[0])                
                self.lineEdit_45.setText(text_stages[0])
 
  ########################################################## @Combobox                      
    def fill_cmbBox(self, cmbBox, select_statement):
        try:
            cmbBox.clear()
            cmbBox.addItem('Select...')
            data = helper.select(select_statement)
            for item in data:
                cmbBox.addItem(item[0])
            cmbBox.setCurrentIndex(0)     
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))        
        
    def fill_cmbBox_status(self):
        self.fill_cmbBox(self.comboBox_5,"Select StatusName from EntityStatus")
        self.fill_cmbBox(self.comboBox,"Select StatusName from EntityStatus")

            
    def fill_cmbBox_adress_type(self):
        self.fill_cmbBox(self.comboBox_15,"Select AddressTypeName from AddressTypes")
        self.fill_cmbBox(self.comboBox_16,"Select AddressTypeName from AddressTypes")
            
    def fill_cmbBox_products(self):
        self.fill_cmbBox(self.comboBox_14,"Select ProductName from Products")
        self.fill_cmbBox(self.comboBox_8,"Select ProductName from Products")

    def fill_cmbBox_products_category(self):
        self.fill_cmbBox(self.comboBox_6,"Select ProductCategoryName from ProductCategory")

    def fill_cmbBox_groups(self):
        self.fill_cmbBox(self.comboBox_13,"Select GroupName from Groups")
        self.fill_cmbBox(self.comboBox_2,"Select GroupName from Groups")
    
    def fill_cmbBox_deals(self):
        self.fill_cmbBox(self.comboBox_12,"Select DealName from deals")
        self.fill_cmbBox(self.comboBox_20,"Select DealName from deals")
        self.fill_cmbBox(self.comboBox_23,"Select DealName from deals")

    def fill_cmbBox_companies(self):
        self.fill_cmbBox(self.comboBox_3,"Select CompanyName from Companies order by CompanyName")
        self.fill_cmbBox(self.comboBox_7,"Select CompanyName from Companies order by CompanyName")
        self.fill_cmbBox(self.comboBox_17,"Select CompanyName from Companies order by CompanyName")
        self.fill_cmbBox(self.comboBox_21,"Select CompanyName from Companies order by CompanyName")
        self.fill_cmbBox(self.comboBox_24,"Select CompanyName from Companies order by CompanyName")

    def fill_cmbBox_contacts(self):
        self.fill_cmbBox(self.comboBox_9,"Select StatusName from Contact order by StatusName")
        self.fill_cmbBox(self.comboBox_10,"Select StatusName from Contact order by StatusName")
        self.fill_cmbBox(self.comboBox_19,"Select StatusName from Contact order by StatusName")
        self.fill_cmbBox(self.comboBox_22,"Select StatusName from Contact order by StatusName")
    
    def fill_cmbBox_stages(self):
        self.fill_cmbBox(self.comboBox_18,"Select StageName from Stages")
        self.comboBox_18.removeItem(0)

    def fill_cmbBox_group(self):
        self.fill_cmbBox(self.comboBox_2,"Select groupname from groups")
        self.fill_cmbBox(self.comboBox_3,"Select groupname from groups")
    def fill_cmbBox_activity_filter(self):
        self.fill_cmbBox(self.comboBox_25,"Select Text from FilterText")
        self.fill_cmbBox(self.comboBox_26,"Select Text from FilterText")
        self.fill_cmbBox(self.comboBox_27,"Select Text from FilterText")

 ########################################################### @delete  
    def delete_deal(self):
        try:
            if self.lineEdit_47.text():
                dealid = self.lineEdit_47.text()
                helper.delete("DELETE FROM deals WHERE dealid ="+dealid)
                self.clear_deal_fields()
                self.fill_deals_qualification_table()
                self.fill_deals_needs_analyse_table()
                self.fill_deals_offer_table()
                self.fill_deals_negotiation_table()
                self.fill_deals_closed_table()
                self.statusBar().showMessage('Deal deleted',5000) 
            else:
                QtWidgets.QMessageBox.information(None, 'No Selection', 'Select one row') 
                
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))    
    def delete_contact(self):
        if self.lineEdit_11.text():
            try:
                contactid = self.lineEdit_11.text()
                helper.delete("DELETE FROM Contact WHERE contactid =" + contactid)
                self.clear_contact_fields()
                self.clear_contact_address_fields()
                self.clear_contact_available_address_table()
                self.delete_contact_address(contactid)
                self.fill_contact_table()
                self.fill_cmbBox_contacts()
                self.statusBar().showMessage('Contact deleted',5000)
            except Exception as e:
                        ErrorLogger.WriteError(traceback.format_exc())
                        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))    
        else:
            QtWidgets.QMessageBox.information(None, 'No Selection', 'Select one row') 

    def delete_contact_address(self,contactid):
        try:
            if contactid: #called from tableWidget
                if helper.select(f"select * from ContactAddressJunction where contactid= {contactid}"):
                    helper.delete(f"""DELETE FROM Addresses
                                  WHERE ROWID IN (
                                      SELECT a.ROWID FROM Addresses a
                                      INNER JOIN ContactAddressJunction b
                                      ON (a.AddressId = b.AddressId )
                                      WHERE b.ContactId = {contactid})""")
                    helper.delete(f"DELETE FROM ContactAddressJunction where contactid= {contactid}")
                    return
            if self.lineEdit_40.text():
                addressid = self.lineEdit_40.text()
                contactid = self.lineEdit_11.text()
                helper.delete("DELETE FROM Addresses WHERE addressid ="+addressid)
                helper.delete(f"DELETE FROM ContactAddressJunction where contactid= {contactid} and addressid = {addressid}")
                self.clear_contact_address_fields()
                self.fill_table_available_addresses()
                self.statusBar().showMessage('Contact Address deleted',5000) 
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))    
                
    def delete_company(self):
        if self.lineEdit_25.text():
            try:
                companyid = self.tableWidget_13.item(self.tableWidget_13.currentRow(),0).text()
                helper.delete(f"""DELETE FROM Addresses WHERE ROWID IN (
                                  SELECT a.ROWID FROM Addresses a
                                  INNER JOIN Companies b
                                  ON (a.AddressId = b.AddressId )
                                  WHERE b.CompanyId = {companyid})""")                
                helper.delete("DELETE FROM Companies WHERE companyid =" + companyid)
                self.clear_company_fields()
                self.clear_company_address_fields()
                self.fill_company_table()
                self.fill_cmbBox_companies()
                self.statusBar().showMessage('Company deleted',5000) 
            except Exception as e:
                        ErrorLogger.WriteError(traceback.format_exc())
                        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
        else:
            QtWidgets.QMessageBox.information(None, 'No Selection', 'Select one row') 
                    
    def delete_product(self):
        try:
            if self.lineEdit_34.text():
                productid = self.lineEdit_34.text()
                helper.delete("DELETE FROM Products WHERE productid =" + productid)
                self.clear_product_fields()
                self.fill_products_table()
                self.fill_cmbBox_products()
                self.statusBar().showMessage('Product deleted',5000) 
            else:
                QtWidgets.QMessageBox.information(None, 'No Selection', 'Select one row') 
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
        
    def delete_task(self):
        if self.lineEdit_8.text():
            try:
                taskid = self.lineEdit_8.text()
                helper.delete("DELETE FROM tasks WHERE taskid =" + taskid)
                self.clear_task_fields()
                self.fill_tasks_table()
                self.statusBar().showMessage('Task deleted',5000) 
            except Exception as e:
                        ErrorLogger.WriteError(traceback.format_exc())
                        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))    
        else:
            QtWidgets.QMessageBox.information(None, 'No Selection', 'Select one row')    
    def delete_event(self):
        if self.lineEdit_17.text():
            try:
                eventid = self.lineEdit_17.text()
                helper.delete("DELETE FROM Events WHERE eventid =" + eventid)
                self.clear_event_fields()
                self.fill_events_table()
                self.statusBar().showMessage('Event deleted',5000) 
            except Exception as e:
                        ErrorLogger.WriteError(traceback.format_exc())
                        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
        else:
            QtWidgets.QMessageBox.information(None, 'No Selection', 'Select one row')
                        
    def delete_call(self):
        if self.lineEdit_22.text():
            try:
                callid = self.lineEdit_22.text()
                helper.delete("DELETE FROM Calls WHERE callid =" + callid)
                self.clear_call_fields()
                self.fill_calls_table()
                self.statusBar().showMessage('Call deleted',5000) 
            except Exception as e:
                        ErrorLogger.WriteError(traceback.format_exc())
                        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
        else:
            QtWidgets.QMessageBox.information(None, 'No Selection', 'Select one row')
            
    def delete_group(self):
        if self.lineEdit_19.text():
            try:
                groupid = self.lineEdit_19.text()
                helper.delete("DELETE FROM groups WHERE groupid =" + groupid)
                self.fill_groups_table()
                self.fill_cmbBox_groups() 
                self.clear_group_fields()
                self.statusBar().showMessage('Group deleted',5000)        
            except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
        else:
            QtWidgets.QMessageBox.information(None, 'No Selection', 'Select one row')
            
    def delete_status(self):
        if self.lineEdit_23.text():
            try:
                statusid = self.lineEdit_23.text()
                helper.delete("DELETE FROM EntityStatus WHERE Statusid =" + statusid)
                self.fill_status_table()
                self.fill_cmbBox_status()
                self.clear_status_fields()
                self.statusBar().showMessage('Status deleted',5000)      
            except Exception as e:
                        ErrorLogger.WriteError(traceback.format_exc())
                        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
        else:
            QtWidgets.QMessageBox.information(None, 'No Selection', 'Select one row')
            
    def delete_product_category(self):
        if self.lineEdit_53.text():
            try:
                procatid = self.lineEdit_53.text()
                helper.delete("DELETE FROM productCategory WHERE productCategoryid =" + procatid)
                self.fill_product_category_table()
                self.fill_cmbBox_products_category()
                self.clear_product_category_fields()
                self.statusBar().showMessage('Product Category deleted',5000)                
            except Exception as e:
                        ErrorLogger.WriteError(traceback.format_exc())
                        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
        else:
            QtWidgets.QMessageBox.information(None, 'No Selection', 'Select one row')
                        
    ######################################################### @clear
    
    def clear_group_fields(self):
        try:
            self.lineEdit_19.clear()
            self.lineEdit_13.clear()
            self.lineEdit_12.clear()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))            
    def clear_task_fields(self):
        try:
            self.lineEdit_8.clear()
            self.lineEdit_7.clear()
            self.comboBox_10.setCurrentIndex(0)
            self.comboBox_12.setCurrentIndex(0)
            self.comboBox_17.setCurrentIndex(0)            
            self.checkBox.setChecked(False)
            self.checkBox_2.setChecked(False)
            self.plainTextEdit_2.clear()
            self.set_date_time()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def clear_contact_fields(self):
        try:
            self.lineEdit_11.clear()
            self.lineEdit_4.clear()
            self.lineEdit_2.clear()
            self.lineEdit_39.clear()
            self.lineEdit_3.clear()
            self.lineEdit_5.clear()
            self.lineEdit_6.clear()
            self.comboBox_13.setCurrentIndex(0)
            self.comboBox_3.setCurrentIndex(0)
            self.comboBox_14.setCurrentIndex(0)
            self.comboBox_5.setCurrentIndex(0)
            self.plainTextEdit.clear()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))    
    def clear_deal_fields(self):
        try:
            self.lineEdit_47.clear()
            self.lineEdit_48.clear()
            self.comboBox_9.setCurrentIndex(0)
            self.comboBox_8.setCurrentIndex(0)
            self.comboBox_18.setCurrentIndex(0)
            self.comboBox_7.setCurrentIndex(0)
            self.lineEdit_57.clear()
            self.dateEdit_3.setDateTime(QDateTime.currentDateTime())
            self.plainTextEdit_7.clear()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def clear_event_fields(self):
        try:
            self.lineEdit_17.clear()
            self.lineEdit_21.clear()
            self.dateEdit_2.setDate(QDate.currentDate())
            self.dateEdit_6.setDate(QDate.currentDate())
            self.lineEdit_38.clear()
            self.comboBox_19.setCurrentIndex(0)
            self.comboBox_20.setCurrentIndex(0)
            self.comboBox_21.setCurrentIndex(0)
            self.lineEdit_42.clear()
            self.plainTextEdit_3.clear()
            self.set_date_time()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def clear_call_fields(self):
        try:
            self.lineEdit_22.clear()
            self.lineEdit_46.clear()
            self.dateEdit_4.setDate(QDate.currentDate())
            self.comboBox_11.setCurrentIndex(0)
            self.comboBox_22.setCurrentIndex(0)
            self.comboBox_23.setCurrentIndex(0)
            self.comboBox_24.setCurrentIndex(0)
            self.plainTextEdit_4.clear()
            self.set_date_time()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def clear_contact_address_fields(self):
        try:
            self.lineEdit_40.clear()
            self.comboBox_15.setCurrentIndex(0)
            self.lineEdit_16.clear()
            self.lineEdit_15.clear()
            self.lineEdit_18.clear()
            self.lineEdit_14.clear()
            self.lineEdit_20.clear()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def clear_contact_available_address_table(self):
        try:
            self.tableWidget_3.setRowCount(0)       
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def clear_contact_activity_table(self):
        try:
            self.tableWidget_17.setRowCount(0)       
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def clear_company_activity_table(self):
        try:
            self.tableWidget_16.setRowCount(0)       
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def clear_company_fields(self):
        try:
            self.lineEdit_25.clear()
            self.lineEdit_26.clear()
            self.lineEdit_27.clear()
            self.lineEdit_28.clear()
            self.plainTextEdit_5.clear()
            self.plainTextEdit_8.clear()
            self.clear_company_address_fields()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def clear_product_fields(self):
        try:
            self.lineEdit_34.clear()
            self.lineEdit_35.clear() 
            self.lineEdit_36.clear()
            self.comboBox_6.setCurrentIndex(0)
            self.lineEdit_37.clear()
            self.plainTextEdit_6.clear()
            self.checkBox_3.setChecked(False)
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def clear_company_address_fields(self):
        try:
            self.lineEdit_41.clear()
            self.comboBox_16.setCurrentIndex(0)
            self.lineEdit_32.clear()
            self.lineEdit_31.clear()
            self.lineEdit_33.clear()
            self.lineEdit_30.clear()
            self.lineEdit_29.clear()         
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def clear_status_fields(self):
        try:
            self.lineEdit_23.clear()
            self.lineEdit_24.clear()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def clear_product_category_fields(self):
        try:
            self.lineEdit_53.clear()
            self.lineEdit_52.clear()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
##################################################### @get
    def get_dealID(self, nr):
        try:
            if nr == 4:
                dealId = self.tableWidget_4.item(self.tableWidget_4.currentRow(),0).text()
            if nr == 6:
                dealId = self.tableWidget_6.item(self.tableWidget_6.currentRow(),0).text()
            elif nr == 7:
                dealId = self.tableWidget_7.item(self.tableWidget_7.currentRow(),0).text()
            elif nr == 8:
                dealId = self.tableWidget_8.item(self.tableWidget_8.currentRow(),0).text()
            elif nr == 9:
                dealId = self.tableWidget_9.item(self.tableWidget_9.currentRow(),0).text()
            self.move_table_fields_deal(dealId)
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
#################################################### @move   
    def move_table_fields_deal(self, dealId):
        try:
            data = helper.select(f"select * from deals where dealid = {dealId}")
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    if column == 0 : self.lineEdit_47.setText(str(item))
                    elif column == 1 : self.lineEdit_48.setText(str(item))
                    elif column == 2 and item:
                        cmbBoxTextT = helper.select("select StatusName from Contact where ContactId="+str(item))
                        if cmbBoxTextT:
                            cmbBoxText = cmbBoxTextT[0][0]
                            self.comboBox_9.setCurrentText(str(cmbBoxText))
                            cmbBoxText = None
                    elif column == 3 and item: 
                        cmbBoxTextT = helper.select("select CompanyName from Companies where CompanyId="+str(item))
                        if cmbBoxTextT:
                            cmbBoxText = cmbBoxTextT[0][0]
                            self.comboBox_7.setCurrentText(str(cmbBoxText))
                            cmbBoxText = None
                    elif column == 4 and item:
                        cmbBoxTextT = helper.select("select ProductName from Products where ProductId="+str(item))
                        if cmbBoxTextT:
                            cmbBoxText = cmbBoxTextT[0][0]
                            self.comboBox_8.setCurrentText(str(cmbBoxText))
                            cmbBoxText = None
                    elif column == 5 and item: 
                        cmbBoxTextT = helper.select("select StageName from Stages where StageID="+str(item))
                        if cmbBoxTextT:
                            cmbBoxText = cmbBoxTextT[0][0]
                            self.comboBox_18.setCurrentText(str(cmbBoxText))
                    elif column == 6 : self.lineEdit_57.setText(str(item))
                    elif column == 7 : 
                        dateD = parse(item, dayfirst=True) 
                        self.dateEdit_3.setDate(dateD)
                    elif column == 8 : self.plainTextEdit_7.setPlainText(str(item))    
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))        
    def move_table_fields_address(self):
        try:
            self.clear_contact_address_fields()
            if self.tableWidget_3.item(self.tableWidget_3.currentRow(),0):
                self.lineEdit_40.setText(self.tableWidget_3.item(self.tableWidget_3.currentRow(),0).text()) 
            if self.tableWidget_3.item(self.tableWidget_3.currentRow(),1):
                self.comboBox_15.setCurrentText(self.tableWidget_3.item(self.tableWidget_3.currentRow(),1).text())
            if self.tableWidget_3.item(self.tableWidget_3.currentRow(),2): 
                self.lineEdit_16.setText(self.tableWidget_3.item(self.tableWidget_3.currentRow(),2).text()) 
            if self.tableWidget_3.item(self.tableWidget_3.currentRow(),3):
                self.lineEdit_15.setText(self.tableWidget_3.item(self.tableWidget_3.currentRow(),3).text()) 
            if self.tableWidget_3.item(self.tableWidget_3.currentRow(),4):    
                self.lineEdit_18.setText(self.tableWidget_3.item(self.tableWidget_3.currentRow(),4).text())
            if self.tableWidget_3.item(self.tableWidget_3.currentRow(),5):
                self.lineEdit_14.setText(self.tableWidget_3.item(self.tableWidget_3.currentRow(),5).text())
            if self.tableWidget_3.item(self.tableWidget_3.currentRow(),6):
                self.lineEdit_20.setText(self.tableWidget_3.item(self.tableWidget_3.currentRow(),6).text())     
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))    
    def move_table_fields_contact(self): # Contact
        try:
            self.clear_contact_fields()
            self.clear_contact_address_fields()
            if self.tableWidget.item(self.tableWidget.currentRow(),0):
                self.lineEdit_11.setText(self.tableWidget.item(self.tableWidget.currentRow(),0).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),1):
                self.lineEdit_4.setText(self.tableWidget.item(self.tableWidget.currentRow(),1).text()) 
            if self.tableWidget.item(self.tableWidget.currentRow(),2):
                self.lineEdit_2.setText(self.tableWidget.item(self.tableWidget.currentRow(),2).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),3):
                self.lineEdit_39.setText(self.tableWidget.item(self.tableWidget.currentRow(),3).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),4):
                self.lineEdit_3.setText(self.tableWidget.item(self.tableWidget.currentRow(),4).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),5):
                self.lineEdit_5.setText(self.tableWidget.item(self.tableWidget.currentRow(),5).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),6):
                self.lineEdit_6.setText(self.tableWidget.item(self.tableWidget.currentRow(),6).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),7):
                self.comboBox_13.setCurrentText(self.tableWidget.item(self.tableWidget.currentRow(),7).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),8):
                self.comboBox_3.setCurrentText(self.tableWidget.item(self.tableWidget.currentRow(),8).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),9):
                self.comboBox_14.setCurrentText(self.tableWidget.item(self.tableWidget.currentRow(),9).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),10):
                self.comboBox_5.setCurrentText(self.tableWidget.item(self.tableWidget.currentRow(),10).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),11):
                self.plainTextEdit.setPlainText(self.tableWidget.item(self.tableWidget.currentRow(),11).text()) 
            if self.tableWidget.item(self.tableWidget.currentRow(),12):
                self.lineEdit_40.setText(self.tableWidget.item(self.tableWidget.currentRow(),12).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),13):
                self.comboBox_15.setCurrentText(self.tableWidget.item(self.tableWidget.currentRow(),13).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),14):
                self.lineEdit_16.setText(self.tableWidget.item(self.tableWidget.currentRow(),14).text()) 
            if self.tableWidget.item(self.tableWidget.currentRow(),15):
                self.lineEdit_15.setText(self.tableWidget.item(self.tableWidget.currentRow(),15).text()) 
            if self.tableWidget.item(self.tableWidget.currentRow(),16):
                self.lineEdit_18.setText(self.tableWidget.item(self.tableWidget.currentRow(),16).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),17):
                self.lineEdit_14.setText(self.tableWidget.item(self.tableWidget.currentRow(),17).text())
            if self.tableWidget.item(self.tableWidget.currentRow(),18):
                self.lineEdit_20.setText(self.tableWidget.item(self.tableWidget.currentRow(),18).text())
            
            self.fill_table_available_addresses()
            self.fill_contact_activity_table()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))    
    def move_table_fields_groups(self): #Groups 
        try:
            self.clear_group_fields()
            if self.tableWidget_2.item(self.tableWidget_2.currentRow(),0):
                self.lineEdit_19.setText(self.tableWidget_2.item(self.tableWidget_2.currentRow(),0).text()) 
            if self.tableWidget_2.item(self.tableWidget_2.currentRow(),1):
                self.lineEdit_13.setText(self.tableWidget_2.item(self.tableWidget_2.currentRow(),1).text()) 
            if self.tableWidget_2.item(self.tableWidget_2.currentRow(),2):
                self.lineEdit_12.setText(self.tableWidget_2.item(self.tableWidget_2.currentRow(),2).text()) 
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def move_table_fields_status(self): # Status 
        try:
            self.clear_status_fields()
            if self.tableWidget_5.item(self.tableWidget_5.currentRow(),0):
                self.lineEdit_23.setText(self.tableWidget_5.item(self.tableWidget_5.currentRow(),0).text()) 
            if self.tableWidget_5.item(self.tableWidget_5.currentRow(),1):
                self.lineEdit_24.setText(self.tableWidget_5.item(self.tableWidget_5.currentRow(),1).text()) 
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))        
    def move_table_fields_company(self): # Company and Address
        try:
            self.clear_company_fields()
            self.clear_company_address_fields()
            self.lineEdit_25.setText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),0).text())
            
            if self.tableWidget_13.item(self.tableWidget_13.currentRow(),1): #Name
                self.lineEdit_26.setText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),1).text()) 
            
            if self.tableWidget_13.item(self.tableWidget_13.currentRow(),3): #Phone
                self.lineEdit_27.setText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),3).text()) 
            
            if self.tableWidget_13.item(self.tableWidget_13.currentRow(),4): #Website
                self.lineEdit_28.setText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),4).text()) 
            
            if self.tableWidget_13.item(self.tableWidget_13.currentRow(),2): #Merged Address
                self.plainTextEdit_8.setPlainText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),2).text()) 
            
            if self.tableWidget_13.item(self.tableWidget_13.currentRow(),5): #Notes/Description
                self.plainTextEdit_5.setPlainText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),5).text()) 
            
            if self.tableWidget_13.item(self.tableWidget_13.currentRow(),6):    
                self.lineEdit_41.setText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),6).text())
            if self.tableWidget_13.item(self.tableWidget_13.currentRow(),7):            
                self.comboBox_16.setCurrentText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),7).text())
            if self.tableWidget_13.item(self.tableWidget_13.currentRow(),8):    
                self.lineEdit_32.setText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),8).text()) 
            if self.tableWidget_13.item(self.tableWidget_13.currentRow(),9):    
                self.lineEdit_31.setText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),9).text()) 
            if self.tableWidget_13.item(self.tableWidget_13.currentRow(),10):    
                self.lineEdit_33.setText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),10).text())
            if self.tableWidget_13.item(self.tableWidget_13.currentRow(),11):    
                self.lineEdit_30.setText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),11).text())
            if self.tableWidget_13.item(self.tableWidget_13.currentRow(),12):    
                self.lineEdit_29.setText(self.tableWidget_13.item(self.tableWidget_13.currentRow(),12).text())
            
            self.fill_company_activity_table()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))            
    def move_table_fields_product(self): 
        try:
            self.clear_product_fields()
            if self.tableWidget_14.item(self.tableWidget_14.currentRow(),0):
                self.lineEdit_34.setText(self.tableWidget_14.item(self.tableWidget_14.currentRow(),0).text()) 
            if self.tableWidget_14.item(self.tableWidget_14.currentRow(),1):
                self.lineEdit_35.setText(self.tableWidget_14.item(self.tableWidget_14.currentRow(),1).text()) 
            if self.tableWidget_14.item(self.tableWidget_14.currentRow(),2):
                self.lineEdit_36.setText(self.tableWidget_14.item(self.tableWidget_14.currentRow(),2).text())
            if self.tableWidget_14.item(self.tableWidget_14.currentRow(),3):
                self.comboBox_6.setCurrentText(self.tableWidget_14.item(self.tableWidget_14.currentRow(),3).text())
            if self.tableWidget_14.item(self.tableWidget_14.currentRow(),4):
                self.lineEdit_37.setText(self.tableWidget_14.item(self.tableWidget_14.currentRow(),4).text())
            if self.tableWidget_14.item(self.tableWidget_14.currentRow(),5):
                self.plainTextEdit_6.setPlainText(self.tableWidget_14.item(self.tableWidget_14.currentRow(),5).text())
            if self.tableWidget_14.item(self.tableWidget_14.currentRow(),6):
                self.checkBox_3.setChecked(True if self.tableWidget_14.item(self.tableWidget_14.currentRow(),6).icon_name == "check" else False) 
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))        
    def move_table_fields_task(self):
        try:
            self.clear_task_fields()
            if (self.tableWidget_10.item(self.tableWidget_10.currentRow(),0)): #id
                self.lineEdit_8.setText(self.tableWidget_10.item(self.tableWidget_10.currentRow(),0).text()) 
            if (self.tableWidget_10.item(self.tableWidget_10.currentRow(),1)): #task name
                self.lineEdit_7.setText(self.tableWidget_10.item(self.tableWidget_10.currentRow(),1).text()) 
            
            if (self.tableWidget_10.item(self.tableWidget_10.currentRow(),2)): #DueDate
                dateF = self.tableWidget_10.item(self.tableWidget_10.currentRow(),2).text() #a string
                dateD = parse(dateF, dayfirst=True)
                self.dateEdit.setDate(dateD) 
            
            if (self.tableWidget_10.item(self.tableWidget_10.currentRow(),3)):#RelatedToContact
                self.comboBox_10.setCurrentText(self.tableWidget_10.item(self.tableWidget_10.currentRow(),3).text()) 
            if (self.tableWidget_10.item(self.tableWidget_10.currentRow(),4)):#RelatedToDeal
                self.comboBox_12.setCurrentText(self.tableWidget_10.item(self.tableWidget_10.currentRow(),4).text()) 
            if (self.tableWidget_10.item(self.tableWidget_10.currentRow(),5)):#RelatedToCompany
                self.comboBox_17.setCurrentText(self.tableWidget_10.item(self.tableWidget_10.currentRow(),5).text()) 
            
            if (self.tableWidget_10.item(self.tableWidget_10.currentRow(),6)):
                self.checkBox.setChecked(True if self.tableWidget_10.item(self.tableWidget_10.currentRow(),6).icon_name == "check" else False) 
            
            if (self.tableWidget_10.item(self.tableWidget_10.currentRow(),7)):
                self.checkBox_2.setChecked(True if self.tableWidget_10.item(self.tableWidget_10.currentRow(),7).icon_name == "check" else False) 
            
            if (self.tableWidget_10.item(self.tableWidget_10.currentRow(),8)):
                self.plainTextEdit_2.setPlainText(self.tableWidget_10.item(self.tableWidget_10.currentRow(),8).text()) 
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def move_table_fields_event(self):
        try:
            self.clear_event_fields()
            self.lineEdit_17.setText(self.tableWidget_11.item(self.tableWidget_11.currentRow(),0).text()) 
            if self.tableWidget_11.item(self.tableWidget_11.currentRow(),1):#Title
                self.lineEdit_21.setText(self.tableWidget_11.item(self.tableWidget_11.currentRow(),1).text()) 
            
            if self.tableWidget_11.item(self.tableWidget_11.currentRow(),2):#DateFrom
                dateF = self.tableWidget_11.item(self.tableWidget_11.currentRow(),2).text() #a string
                dateD = parse(dateF, dayfirst=True) 
                self.dateEdit_6.setDate(dateD)
            
            if self.tableWidget_11.item(self.tableWidget_11.currentRow(),3):#TimeFrom
                dateF = self.tableWidget_11.item(self.tableWidget_11.currentRow(),3).text()
                dateD = parse(dateF) 
                dateK = dateD.time()
                self.timeEdit_2.setTime(dateK)
            
            
            if self.tableWidget_11.item(self.tableWidget_11.currentRow(),4):#DateTo
                dateF = self.tableWidget_11.item(self.tableWidget_11.currentRow(),4).text() #a string
                dateD = parse(dateF, dayfirst=True) 
                self.dateEdit_2.setDate(dateD)
            
            if self.tableWidget_11.item(self.tableWidget_11.currentRow(),5):#TimeTo
                dateF = self.tableWidget_11.item(self.tableWidget_11.currentRow(),5).text()
                dateD = parse(dateF) 
                dateK = dateD.time()
                self.timeEdit_3.setTime(dateK)
            
            if self.tableWidget_11.item(self.tableWidget_11.currentRow(),6):#Location
                self.lineEdit_38.setText(self.tableWidget_11.item(self.tableWidget_11.currentRow(),6).text()) 
            
            if (self.tableWidget_11.item(self.tableWidget_11.currentRow(),7)):#RelatedToContact
                self.comboBox_19.setCurrentText(self.tableWidget_11.item(self.tableWidget_11.currentRow(),7).text()) 
            
            if (self.tableWidget_11.item(self.tableWidget_11.currentRow(),8)):#RelatedToDeal
                self.comboBox_20.setCurrentText(self.tableWidget_11.item(self.tableWidget_11.currentRow(),8).text()) 
            
            if (self.tableWidget_11.item(self.tableWidget_11.currentRow(),9)):#RelatedToCompany
                self.comboBox_21.setCurrentText(self.tableWidget_11.item(self.tableWidget_11.currentRow(),9).text()) 

            if self.tableWidget_11.item(self.tableWidget_11.currentRow(),10):#Participants
                self.lineEdit_42.setText(self.tableWidget_11.item(self.tableWidget_11.currentRow(),10).text())
            if self.tableWidget_11.item(self.tableWidget_11.currentRow(),11):#Description
                self.plainTextEdit_3.setPlainText(self.tableWidget_11.item(self.tableWidget_11.currentRow(),11).text()) 
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def move_table_fields_call(self):
        try:
            self.clear_call_fields()
            if self.tableWidget_12.item(self.tableWidget_12.currentRow(),0):#Id
                self.lineEdit_22.setText(self.tableWidget_12.item(self.tableWidget_12.currentRow(),0).text()) 
            if self.tableWidget_12.item(self.tableWidget_12.currentRow(),1):#ToFrom
                self.lineEdit_46.setText(self.tableWidget_12.item(self.tableWidget_12.currentRow(),1).text()) 
            
            if self.tableWidget_12.item(self.tableWidget_12.currentRow(),2):#startTime
                dateF = self.tableWidget_12.item(self.tableWidget_12.currentRow(),2).text()
                dateD = parse(dateF) 
                dateK = dateD.time()
                self.timeEdit.setTime(dateK)
            
            if self.tableWidget_12.item(self.tableWidget_12.currentRow(),3):#CallDate
                dateF = self.tableWidget_12.item(self.tableWidget_12.currentRow(),3).text() #a string
                dateD = parse(dateF, dayfirst=True) 
                self.dateEdit_4.setDate(dateD)
            
            if self.tableWidget_12.item(self.tableWidget_12.currentRow(),4):#CallType
                self.comboBox_11.setCurrentText(self.tableWidget_12.item(self.tableWidget_12.currentRow(),4).text())
            
            if (self.tableWidget_12.item(self.tableWidget_12.currentRow(),5)):#RelatedToContact
                self.comboBox_22.setCurrentText(self.tableWidget_12.item(self.tableWidget_12.currentRow(),5).text())
            
            if (self.tableWidget_12.item(self.tableWidget_12.currentRow(),6)):#RelatedToDeal
                self.comboBox_23.setCurrentText(self.tableWidget_12.item(self.tableWidget_12.currentRow(),6).text()) 
            
            if (self.tableWidget_12.item(self.tableWidget_12.currentRow(),7)):#RelatedToCompany
                self.comboBox_24.setCurrentText(self.tableWidget_12.item(self.tableWidget_12.currentRow(),7).text()) 

            if self.tableWidget_12.item(self.tableWidget_12.currentRow(),8):
                self.plainTextEdit_4.setPlainText(self.tableWidget_12.item(self.tableWidget_12.currentRow(),8).text()) 
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def move_table_fields_product_category(self): # Product Category 
        try:
            self.clear_product_category_fields()
            if self.tableWidget_15.item(self.tableWidget_15.currentRow(),0):
                self.lineEdit_53.setText(self.tableWidget_15.item(self.tableWidget_15.currentRow(),0).text()) 
            if self.tableWidget_15.item(self.tableWidget_15.currentRow(),1):
                self.lineEdit_52.setText(self.tableWidget_15.item(self.tableWidget_15.currentRow(),1).text())         
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
####################################################### @save 
    def save_contact(self):
        try:
            fn = self.lineEdit_4.text()
            ln = self.lineEdit_2.text()
            st_n = fn + ' ' + ln
            ti = self.lineEdit_39.text()
            em = self.lineEdit_3.text()
            mo = self.lineEdit_5.text()
            ph = self.lineEdit_6.text()
            if self.comboBox_13.currentIndex() != 0:
                gName = self.comboBox_13.currentText()
                gr = helper.select(f"select GroupID from Groups where GroupName like '{gName}'")[0][0]
            else:
                gr = 0
            if self.comboBox_3.currentIndex() != 0:
                cName = self.comboBox_3.currentText()
                co = helper.select(f"select CompanyID from Companies where CompanyName like '{cName}'")[0][0]
            else:
                co = 0
            if self.comboBox_14.currentIndex() != 0:
                prName = self.comboBox_14.currentText()
                pr = helper.select(f"select ProductId from Products where ProductName like '{prName}'")[0][0]
            else:
                pr = 0
            if self.comboBox_5.currentIndex() != 0:
                stName = self.comboBox_5.currentText()
                st = helper.select(f"select StatusId from EntityStatus where StatusName like '{stName}'")[0][0]
            else:
                st = 0
            
            cn = str(self.plainTextEdit.toPlainText())
           
            if self.comboBox_15.currentIndex() != 0:
                adName = self.comboBox_15.currentText()
                address_type = helper.select(f"select AddressTypeId from AddressTypes where AddressTypeName like '{adName}'")[0][0]
            else:
                address_type = 0            
                
            street = self.lineEdit_16.text()
            city = self.lineEdit_15.text()
            state = self.lineEdit_18.text()
            country = self.lineEdit_14.text()
            zip_code = self.lineEdit_20.text()
    
            
            contact = (fn, ln, ti, em, mo, ph, pr, co, gr, st, cn, st_n)
            address = (address_type, street, city, state, country, zip_code)
            contactid = self.lineEdit_11.text()
            addressID = self.lineEdit_40.text()  
            
            if self.lineEdit_11.text():
                helper.edit("UPDATE contact SET FirstName=?, LastName=?, Title=?, Email=?, Mobile=?, Phone=?, ProductID=?, CompanyID=?, GroupID=?, EntityStatusID=?, ContactDescription=?, StatusName=? WHERE ContactID="+contactid,contact)                                                                                                                                           
                if self.lineEdit_40.text():
                    helper.edit("UPDATE Addresses SET AddressTypeID=?, MailingStreet=?, MailingCity=?, MailingState=?, MailingCountry=?, MailingZip=? WHERE AddressID="+addressID,address)
                else:
                    helper.insert("INSERT INTO Addresses (AddressTypeID, MailingStreet, MailingCity, MailingState, MailingCountry, MailingZip) Values(?,?,?,?,?,?)",address)
                    newAddressID = str(helper.select("select max(AddressID) from Addresses")[0][0])
                    newJunction = (contactid, newAddressID)
                    helper.insert("INSERT INTO ContactAddressJunction (ContactID, AddressID) Values(?,?)",newJunction)
                    self.lineEdit_40.setText(newAddressID)
            else:  # every new contact gets a addresId no matter if he has a address 
                helper.insert("""INSERT INTO Contact (FirstName, LastName, Title, Email, Mobile, 
                            Phone, ProductID, CompanyID, GroupID, EntityStatusID, ContactDescription, StatusName) 
                            Values(?,?,?,?,?,?,?,?,?,?,?,?)""",contact)
                helper.insert("INSERT INTO Addresses (AddressTypeID, MailingStreet, MailingCity, MailingState, MailingCountry, MailingZip) Values(?,?,?,?,?,?)",address)
                
                newContactID = str(helper.select("select max(ContactID) from Contact")[0][0])
                newAddressID = str(helper.select("select max(AddressID) from Addresses")[0][0])
                newJunction = (newContactID, newAddressID)
                helper.insert("INSERT INTO ContactAddressJunction (ContactID, AddressID) Values(?,?)",newJunction)
                self.lineEdit_11.setText(newContactID)
                self.lineEdit_40.setText(newAddressID)
            self.statusBar().showMessage('Contact saved',5000)                
            self.clear_contact_available_address_table()
            self.fill_contact_table()
            self.fill_cmbBox_contacts()
            self.fill_table_available_addresses()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))        
    def save_call(self):
        try:
            callid = self.lineEdit_22.text()
            to_from = self.lineEdit_46.text()
            call_date = self.dateEdit_4.date().toPyDate()
            start_time = self.timeEdit.time().toString(self.timeEdit.displayFormat())
            call_type = self.comboBox_11.currentText()
            if self.comboBox_22.currentIndex() != 0: # Contact
                contactName = self.comboBox_22.currentText()
                related_to_contact = helper.select(f"select ContactID from Contact where StatusName like '{contactName}'")[0][0]
            else:
                related_to_contact = 0
            
            if self.comboBox_23.currentIndex() != 0: # Deal
                DealName = self.comboBox_23.currentText()
                related_to_deal = helper.select(f"select DealId from Deals where DealName like '{DealName}'")[0][0]
            else:
                related_to_deal = 0
            
            if self.comboBox_24.currentIndex() != 0: # Company
                CompanyName = self.comboBox_24.currentText()
                related_to_company = helper.select(f"select CompanyId from Companies where CompanyName like '{CompanyName}'")[0][0]
            else:
                related_to_company = 0
            
            agenda = str(self.plainTextEdit_4.toPlainText())
    
            call = (to_from, start_time, call_date, call_type, related_to_contact, related_to_deal, related_to_company, agenda)
            if self.lineEdit_22.text():
                helper.edit("UPDATE calls SET ToFrom=?,  CallTime=?, calldate=?, CallType=?, RelatedToContact=?, RelatedToDeal=?, RelatedToCompany=?, CallAgenda=? WHERE CallID="+callid,call)
            else:      
                helper.insert("INSERT INTO Calls (ToFrom, CallTime, CallDate, CallType, RelatedToContact, RelatedToDeal, RelatedToCompany, CallAgenda) Values(?,?,?,?,?,?,?,?)",call)
                newCallID = str(helper.select("select max(CallID) from Calls")[0][0])
                self.lineEdit_22.setText(newCallID)
            self.fill_calls_table()
            self.fill_contact_activity_table()
            self.fill_company_activity_table()            
            self.statusBar().showMessage('Call saved',5000)                
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))       
    def save_task(self):
        try:
            taskid = self.lineEdit_8.text()
            task_name = self.lineEdit_7.text()
            due_date = self.dateEdit.date().toPyDate()
            
            if self.comboBox_10.currentIndex() != 0: # Contact
                contactName = self.comboBox_10.currentText()
                related_to_contact = helper.select(f"select ContactID from Contact where StatusName like '{contactName}'")[0][0]
            else:
                related_to_contact = 0
            
            if self.comboBox_12.currentIndex() != 0: # Deal
                DealName = self.comboBox_12.currentText()
                related_to_deal = helper.select(f"select DealId from Deals where DealName like '{DealName}'")[0][0]
            else:
                related_to_deal = 0
            
            if self.comboBox_17.currentIndex() != 0: # Company
                CompanyName = self.comboBox_17.currentText()
                related_to_company = helper.select(f"select CompanyId from Companies where CompanyName like '{CompanyName}'")[0][0]
            else:
                related_to_company = 0
            
            high_priority = self.checkBox.isChecked()
            complete = self.checkBox_2.isChecked()
            description = self.plainTextEdit_2.toPlainText()
            
            task = (task_name, due_date, related_to_contact, related_to_deal , related_to_company, high_priority, complete, description)
            
            if self.lineEdit_8.text():
                helper.edit("UPDATE tasks SET TaskName=?, DueDate=?, RelatedToContact=?,RelatedToDeal=?,RelatedToCompany=?, HighPriority=?, Completed=?, Description=? WHERE TaskID="+taskid,task)
            else:    
                helper.insert("INSERT INTO Tasks (TaskName, DueDate, RelatedToContact, RelatedToDeal, RelatedToCompany, HighPriority, Completed , Description) Values(?,?,?,?,?,?,?,?)",task)
                newTaskID = str(helper.select("select max(TaskID) from Tasks")[0][0])
                self.lineEdit_8.setText(newTaskID)
            
            self.statusBar().showMessage('Task saved',5000)
            self.fill_tasks_table()
            self.fill_contact_activity_table()
            self.fill_company_activity_table()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def save_event(self):
        try:
            eventid = self.lineEdit_17.text()
            title = self.lineEdit_21.text()
            from_date = self.dateEdit_6.date().toPyDate() #toString(self.dateEdit_6.displayFormat())
            from_time = self.timeEdit_2.time().toString(self.timeEdit_2.displayFormat())
            to_date =   self.dateEdit_2.date().toPyDate() #toString(self.dateEdit_2.displayFormat())
            to_time = self.timeEdit_3.time().toString(self.timeEdit_3.displayFormat())
            location = self.lineEdit_38.text()
            if self.comboBox_19.currentIndex() != 0: # Contact
                contactName = self.comboBox_19.currentText()
                related_to_contact = helper.select(f"select ContactID from Contact where StatusName like '{contactName}'")[0][0]
            else:
                related_to_contact = 0
            
            if self.comboBox_20.currentIndex() != 0: # Deal
                DealName = self.comboBox_20.currentText()
                related_to_deal = helper.select(f"select DealId from Deals where DealName like '{DealName}'")[0][0]
            else:
                related_to_deal = 0
            
            if self.comboBox_21.currentIndex() != 0: # Company
                CompanyName = self.comboBox_21.currentText()
                related_to_company = helper.select(f"select CompanyId from Companies where CompanyName like '{CompanyName}'")[0][0]
            else:
                related_to_company = 0            

            participants = self.lineEdit_42.text() 
            description = self.plainTextEdit_3.toPlainText()
            event = (title, from_date, from_time, to_date, to_time, location, related_to_contact, related_to_deal, related_to_company, participants, description)
            
            if self.lineEdit_17.text():
                helper.edit("UPDATE Events SET EventTitle=?, DateFrom=?, TimeFrom=?, DateTo=?, TimeTo=?, Location=?, RelatedToContact=?, RelatedToDeal=?, RelatedToCompany=?, Participants=?, Description=? WHERE EventID="+eventid,event)
            else:            
                helper.insert("INSERT INTO Events (EventTitle, DateFrom, TimeFrom, DateTo, TimeTo, Location, RelatedToContact, RelatedToDeal, RelatedToCompany, Participants, Description) Values(?,?,?,?,?,?,?,?,?,?,?)",event)
                newEventID = str(helper.select("select max(EventID) from Events")[0][0])
                self.lineEdit_17.setText(newEventID)
            self.statusBar().showMessage('Event saved',5000)                
            self.fill_events_table()
            self.fill_contact_activity_table()
            self.fill_company_activity_table()
            
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))        
    def save_deal(self):
        try:
            dealID = self.lineEdit_47.text()
            dealName = self.lineEdit_48.text()
             
            if self.comboBox_9.currentIndex() != 0:
                cName = self.comboBox_9.currentText()
                contactNameId = helper.select(f"select ContactId from Contact where StatusName like '{cName}'") 
                contactId = contactNameId[0][0]
                cName = None
            else:
                contactId = 0
            if self.comboBox_7.currentIndex() != 0:
                cName = self.comboBox_7.currentText()
                companyNameId = helper.select(f"select CompanyID from Companies where CompanyName like '{cName}'") 
                companyId = companyNameId[0][0]
                cName = None                
            else:
                companyId = 0
            if self.comboBox_8.currentIndex() != 0:
                cName = self.comboBox_8.currentText()
                productNameId = helper.select(f"select ProductId from Products where ProductName like '{cName}'") 
                productId = productNameId[0][0]
                cName = None                
            else:
                productId = 0
           
            cName = self.comboBox_18.currentText()
            stageNameId = helper.select(f"select StageId from stages where StageName like '{cName}'") 
            stageId = stageNameId[0][0]
            cName = None
            
            if self.lineEdit_57.text():
                amount = float(self.lineEdit_57.text())
            else:
                amount = 0
                
            closingDate = self.dateEdit_3.dateTime().toString(self.dateEdit_3.displayFormat())
            description = self.plainTextEdit_7.toPlainText()
            deal = (dealName, contactId, companyId, productId, stageId, amount, closingDate, description)
            
            if self.lineEdit_47.text():
                helper.edit("UPDATE deals SET dealName=?, ContactId=?, CompanyId=?, ProductId=?, StageId=?, Amount=?, closingDate=?, Description=? WHERE DealID="+dealID,deal)
            else:
                helper.insert("INSERT INTO deals (dealName, ContactId, CompanyId, ProductId, stageId, Amount, closingDate, Description) Values(?,?,?,?,?,?,?,?)",deal)
                newDealID = str(helper.select("select max(DealId) from Deals")[0][0])
                self.lineEdit_47.setText(newDealID)
            self.fill_deals_qualification_table()
            self.fill_deals_needs_analyse_table()
            self.fill_deals_offer_table()
            self.fill_deals_negotiation_table()
            self.fill_deals_closed_table()
            self.fill_cmbBox_deals()
            self.fill_deal_stage_amount()
            self.statusBar().showMessage('Deal saved',5000)
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def save_product(self):
        try:
            productID = self.lineEdit_34.text()
            product_name = self.lineEdit_35.text()
            product_code = self.lineEdit_36.text()
            if self.comboBox_6.currentIndex() != 0:
                prCategoryName = self.comboBox_6.currentText()
                CategoryId = helper.select(f"select ProductCategoryId from ProductCategory where ProductCategoryName like '{prCategoryName}'") 
                result = CategoryId.pop()             
                category = (result[0])
            else:
                category = 0   
           
            if self.lineEdit_37.text():
                unit_price = self.lineEdit_37.text()
            else:
                unit_price = 0 
            if unit_price:
                if (not unit_price.isdigit()):
                        QtWidgets.QMessageBox.critical(None, 'Exception', 'Only numbers in Unit Price')                
                        return
            description = self.plainTextEdit_6.toPlainText()
            prIsActiv = self.checkBox_3.isChecked()
            
            product = (product_name, product_code, category, unit_price, description, prIsActiv)
            if self.lineEdit_34.text():
                helper.edit("UPDATE Products SET ProductName=?, ProductCode=?, ProductCategory=?, UnitPrice=?, Description=?, ProductActive=? WHERE ProductID="+productID,product)
            else:    
                helper.insert("INSERT INTO Products (ProductName, ProductCode, ProductCategory, UnitPrice, Description, ProductActive) Values(?,?,?,?,?,?)",product)
                newProductID = str(helper.select("select max(ProductId) from Products")[0][0])
                self.lineEdit_34.setText(newProductID)
            self.fill_products_table()
            self.fill_cmbBox_products()
            self.statusBar().showMessage('Product saved',5000)
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))  
    def save_company(self):
        try:
            company_name = self.lineEdit_26.text()
            merged_address = self.plainTextEdit_8.toPlainText()
            phone = self.lineEdit_27.text()
            website = self.lineEdit_28.text()
            description = self.plainTextEdit_5.toPlainText()
            
            address_type = self.comboBox_16.currentIndex() 
            street = self.lineEdit_32.text()
            city = self.lineEdit_31.text()
            state = self.lineEdit_33.text()
            country = self.lineEdit_30.text()
            zip_code = self.lineEdit_29.text()
            
            company = (company_name, phone, website, description, merged_address)
            address = (address_type, street, city, state, country, zip_code)
            
            if self.lineEdit_25.text():
                companyid = self.lineEdit_25.text()
                helper.edit("UPDATE companies SET CompanyName=?, Phone=?, Website=?, Description=?, MergedAddress=? WHERE CompanyID="+companyid,company)
                if self.lineEdit_41.text():
                    addressId = self.lineEdit_41.text()
                    helper.edit("UPDATE addresses SET AddressTypeID=?, MailingStreet=?, MailingCity=?, MailingState=?, MailingCountry=?, MailingZip=? WHERE AddressID="+addressId,address)
                else:
                    helper.insert("INSERT INTO Addresses (AddressTypeID, MailingStreet, MailingCity, MailingState, MailingCountry, MailingZip) Values(?,?,?,?,?,?)",address)
                    newAddressID = str(helper.select("select max(AddressID) from Addresses")[0][0])
                    helper.edit("UPDATE Companies SET AddressID=? where CompanyID="+companyid,(newAddressID,))
                    self.lineEdit_41.setText(newAddressID)
            else:
                helper.insert("INSERT INTO Companies (CompanyName, Phone, Website, Description, MergedAddress) Values(?,?,?,?,?)",company)
                helper.insert("INSERT INTO Addresses (AddressTypeID, MailingStreet, MailingCity, MailingState, MailingCountry, MailingZip) Values(?,?,?,?,?,?)",address)
                newCompanyID = str(helper.select("select max(CompanyID) from Companies")[0][0])
                newAddressID = str(helper.select("select max(AddressID) from Addresses")[0][0])
                helper.edit("UPDATE Companies SET AddressID=? where CompanyID="+newCompanyID,(newAddressID,))
                self.lineEdit_25.setText(newCompanyID)
                self.lineEdit_41.setText(newAddressID)

            self.fill_company_table()
            self.fill_cmbBox_companies()
            self.statusBar().showMessage('Company and address saved',5000)           
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))   
    def save_group(self):
        try:
            gn = self.lineEdit_13.text()
            gk = self.lineEdit_12.text()
            
            group = (gn, gk)
            if self.lineEdit_19.text():
                groupid = self.lineEdit_19.text()
                helper.edit("UPDATE groups SET GroupName=?, GroupComments=? WHERE GroupID="+groupid,group)
                self.statusBar().showMessage('Group saved',5000)
            else:
                helper.insert("INSERT INTO groups (GroupName, GroupComments) Values(?,?)",group)
                self.statusBar().showMessage('Group added',5000)
                newGroupID = str(helper.select("select max(GroupID) from Groups")[0][0])
                self.lineEdit_19.setText(newGroupID)
            self.fill_groups_table()
            self.fill_cmbBox_groups()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def save_status(self):
        try:
            sn = self.lineEdit_24.text()
            status = (sn,)
            if self.lineEdit_23.text():
                statusid = self.lineEdit_23.text()
                helper.edit("UPDATE EntityStatus SET StatusName=? WHERE StatusID="+statusid,status)
                self.statusBar().showMessage('Status saved',5000)
            else:
                helper.insert("INSERT INTO EntityStatus (StatusName) Values(?)",status)
                self.statusBar().showMessage('Status added',5000)
                newStatusID = str(helper.select("select max(StatusID) from EntityStatus")[0][0])
                self.lineEdit_23.setText(newStatusID)
            self.fill_status_table()
            self.fill_cmbBox_status()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def save_stages(self):
        try:
            s1 = self.lineEdit_9.text()
            s2 = self.lineEdit_10.text()
            s3 = self.lineEdit_43.text()
            s4 = self.lineEdit_44.text()
            s5 = self.lineEdit_45.text()
            
            stages = (s1,s2,s3,s4,s5)
            
            for i, text_stages in enumerate(stages):
                stageid = str(i + 1)
                helper.edit(f"UPDATE Stages SET StageName=? WHERE StageID="+stageid,(text_stages,))
            
            self.statusBar().showMessage('Status saved',5000)
            self.fill_stages_label()
            self.fill_cmbBox_stages()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def save_product_category(self):
        try:
            pc = self.lineEdit_52.text()
            procat = (pc,)
            if self.lineEdit_53.text():
                procatid = self.lineEdit_53.text()
                helper.edit("UPDATE ProductCategory SET ProductCategoryName=? WHERE ProductCategoryID="+procatid,procat)
                self.statusBar().showMessage('Product Category saved',5000)
            else:
                helper.insert("INSERT INTO ProductCategory (ProductCategoryName) Values(?)",procat)
                self.statusBar().showMessage('Product Category added',5000)
                newProCatID = str(helper.select("select max(ProductCategoryID) from ProductCategory")[0][0])
                self.lineEdit_53.setText(newProCatID)
            self.fill_product_category_table()
            self.fill_cmbBox_products_category()        
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
###################################################### @set   
    def set_date_time(self):
        try:
            self.dateEdit.setDate(QDate.currentDate())
            self.dateEdit_4.setDate(QDate.currentDate())
            self.dateEdit_6.setDate(QDate.currentDate())
            self.dateEdit_2.setDate(QDate.currentDate())
            self.dateEdit_3.setDate(QDate.currentDate())
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))        
###################################################### @filter
    def filter_no_filter(self):
        try:
            data = helper.select("select * from ContactQ")
            self.tableWidget.setRowCount(0)
            for row, form in enumerate(data):
                row_position = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row_position)               
                for column, item in enumerate(form):
                    if item:
                        self.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
            self.comboBox.setCurrentIndex(0)
            self.comboBox_2.setCurrentIndex(0)
            self.lineEdit.clear()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))

    def filter_task(self, index):
        try:
            text = self.comboBox_26.model().itemFromIndex(index)
            text = text.text()
            if text == 'Only Today':
                data = helper.select("select * from taskQ where DueDate = (select date('now'))")
            elif text == 'From Today On':
                data = helper.select("select * from taskQ where DueDate >= (select date('now'))")
            elif text == 'All':
                data = helper.select("select * from taskQ")
            else:
                return
            
            self.tableWidget_10.setRowCount(0)

            if data:
                cb = QCheckBox()
                for row, form in enumerate(data):
                    row_position = self.tableWidget_10.rowCount()
                    self.tableWidget_10.insertRow(row_position)
                    for column, item in enumerate(form):
                        if column in [6, 7]:
                            status_icon = None
                            
                            if item :
                                icon = r"resources/icon/check.ico"
                                status_icon = TheIconItem("check", icon)
                            else:
                                icon = r"resources/icon/uncheck.ico"
                                status_icon = TheIconItem("uncheck", icon)
                            
                            self.tableWidget_10.setItem(row, column, status_icon)
                        else:
                            if item:
                                self.tableWidget_10.setItem(row, column, QTableWidgetItem(str(item)))                

        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def filter_event(self, index):
        try:
            text = self.comboBox_25.model().itemFromIndex(index)
            text = text.text()
            if text == 'Only Today':
                data = helper.select("select * from eventq where DateFrom = (select date('now'))")
            elif text == 'From Today On':
                data = helper.select("select * from eventq where DateFrom >= (select date('now'))")
            elif text == 'All':
                data = helper.select("select * from eventq")
            else:
                return

            self.tableWidget_11.setRowCount(0)

            if data:
                for row, form in enumerate(data):
                    row_position = self.tableWidget_11.rowCount()
                    self.tableWidget_11.insertRow(row_position)            
                    for column, item in enumerate(form):
                        if item:
                            self.tableWidget_11.setItem(row, column, QTableWidgetItem(str(item)))

        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def filter_call(self, index):
        try:
            text = self.comboBox_27.model().itemFromIndex(index)
            text = text.text()
            if text == 'Only Today':
                data = helper.select("select * from callQ where CallDate = (select date('now'))")
            elif text == 'From Today On':
                data = helper.select("select * from callQ where CallDate >= (select date('now'))")
            elif text == 'All':
                data = helper.select("select * from callQ")
            else:
                return

            self.tableWidget_12.setRowCount(0)

            if data:
                for row, form in enumerate(data):
                    row_position = self.tableWidget_12.rowCount()
                    self.tableWidget_12.insertRow(row_position)            
                    for column, item in enumerate(form):
                        if item:
                            self.tableWidget_12.setItem(row, column, QTableWidgetItem(str(item)))

        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def filter_by_group(self, index):
        try:
            text = self.comboBox_2.model().itemFromIndex(index)
            text = text.text()
            data = helper.select(f"select * from ContactQ where GroupName= '{text}'")
            self.tableWidget.setRowCount(0)
            for row, form in enumerate(data):
                row_position = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row_position)            
                for column, item in enumerate(form):
                    self.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1
            self.comboBox.setCurrentIndex(0)
            self.lineEdit.clear()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def filter_by_status(self, index):
        try:
            text = self.comboBox.model().itemFromIndex(index)
            text = text.text()
            data = helper.select(f"select * from ContactQ where StatusName= '{text}'")
            self.tableWidget.setRowCount(0)
            for row, form in enumerate(data):
                row_position = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row_position)            
                for column, item in enumerate(form):
                    self.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1
            self.comboBox_2.setCurrentIndex(0)
            self.lineEdit.clear()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def filter_by_text(self, text):
        try:
            data = helper.select(f"select * from Contact where FirstName LIKE '{text}%' OR LastName LIKE '{text}%'")
            self.tableWidget.setRowCount(0)
            for row, form in enumerate(data):
                row_position = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row_position)              
                for column, item in enumerate(form):
                    self.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1
            self.comboBox.setCurrentIndex(0)
            self.comboBox_2.setCurrentIndex(0)
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))

######################################################        
    def about_CRM(self):
        try:
            self.about_dialog = Ui_About()
            self.about_dialog.show()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def show_window_available(self):
        try:
            self.available_dialog = Ui_available()
            self.available_dialog.show()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def new_database(self):
        try:
            target = QFileDialog.getSaveFileName(self, 'Save File',"Select Folder and give name")
            target = (target[0])
            contact_db_file = r"assets/db/db_copy/Contact.db"
            if target:
                import shutil
    
                original = contact_db_file
                target = target + ".db"
    
                shutil.copyfile(original, target)
                self.open_database(target)
                self.refresh()
                self.statusBar().showMessage(f'New database created in {target}',5000)
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def open_database(self, db_file):
        try:
            global helper
            
            if not db_file:
                options = QFileDialog.Options()
                options |= QFileDialog.DontUseNativeDialog
                db_file, _ = QFileDialog.getOpenFileName(self, "Chose a Database", "",
                                                  "DB File (*.db)", options=options)
            if db_file:
                helper = SqlHelper(db_file)
                cuurent_db_file = open(r"assets/db/current_db.txt", "w")
                cuurent_db_file.write(db_file)
                cuurent_db_file.close()
                self.refresh()
                self.statusBar().showMessage(f'Opened {db_file}',5000)
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
    def export_as_csv(self):
        try:
            folder = "Contact details"
            if not exists(folder):
                mkdir(folder)
            file = "contact_details.csv"
            try:
                with open(folder + "/" + file, "w", newline='') as csv_file:
                    csv_writer = csv.writer(csv_file)
                    len = self.tableWidget.rowCount()
                    if not len:
                        msg_box = QMessageBox()
                        msg_box.setWindowIcon(QtGui.QIcon(r"resources/icon/favicon-32x32.png"))
                        msg_box.setWindowTitle("CRM")
                        msg_box.setText(f"Can not export empty table")
                        msg_box.setIcon(QMessageBox.Critical)
                        msg_box.exec_()
                        return None
                    for row in range(len):
                        if row == 0:
                            header = []
                        csv_row = []
                        for column in range(self.tableWidget.columnCount()):
                            if row == 0:
                                _item = self.tableWidget.horizontalHeaderItem(column)
                                if _item:
                                    item = self.tableWidget.horizontalHeaderItem(column).text()
                                    header.append(item)
                                else:
                                    header.append("")
    
                            _item = self.tableWidget.item(row, column)
                            if _item:
                                item = self.tableWidget.item(row, column).text()
                                csv_row.append(item)
                            else:
                                csv_row.append("")
                        if row == 0:
                            csv_writer.writerow(header)
                        csv_writer.writerow(csv_row)
            except:
                msg_box = QMessageBox()
                msg_box.setWindowIcon(QtGui.QIcon(r"resources/icon/favicon-32x32.png"))
                msg_box.setWindowTitle("CRM")
                msg_box.setText(f"Permission denied:\nClose the file {file} if it is open somewhere.")
                msg_box.setIcon(QMessageBox.Critical)
                msg_box.exec_()
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
        QtWidgets.QMessageBox.information(None, 'Exported', f"Saved into folder {folder}/{file}")

    def show_csv_window_contact(self):
        try:
            self.import_window = Ui_csv_file_import_contact(helper)
            self.import_window.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
            self.import_window.show() 
            self.import_window.pushButtonClose.clicked.connect(self.refresh)
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))        
    def show_csv_window_company(self):
        try:
            self.import_window = Ui_csv_file_import_company(helper)
            self.import_window.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
            self.import_window.show() 
            self.import_window.pushButtonClose.clicked.connect(self.refresh)
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))        

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    make_a_expire_date()
    if os.path.isfile("hanso.txt"):
        if datetime.date(datetime.now()) >= get_expire_date() :
            vef_key, ok = QInputDialog().getText(window, "Expired :-(",
                                              "Your try has expired. Enter your license key:", QLineEdit.Normal)
            if ok and vef_key:
                if not verify_account_key(vef_key):
                    QtWidgets.QMessageBox.critical(None, 'Error','This key is not valid' )
                    sys.exit()
                else:
                    os.remove("hanso.txt")
                    QtWidgets.QMessageBox.information(None, 'Success','Your Software is acitvated, enjoy' )
            else:
                  sys.exit()

    window.show()
    sys.exit(app.exec())
