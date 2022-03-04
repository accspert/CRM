
from PyQt5.QtWidgets import QMessageBox
import sqlite3
class SqlHelper:
    
    def __init__(self,conn_str = None):
        self.conn = None
        self.cursor = None
        
        if conn_str:
            self.open(conn_str)
            
    def open(self, name):
        try:
            self.conn = sqlite3.connect(name) #,  detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
            self.cursor = self.conn.cursor()
        except sqlite3.Error as e:
            return

    def edit(self, query, updates):  # Update
        c = self.cursor
        c.execute(query, updates)
        self.conn.commit()
    def delete(self, query):         # Delete
        c = self.cursor
        c.execute(query)
        self.conn.commit()
    def insert(self, query, inserts):  # Insert with parameter
        c = self.cursor
        c.execute(query, inserts)
        self.conn.commit()
    def select(self, query):        # Select *
        c = self.cursor
        c.execute(query)
        return c.fetchall()
    def select_para(self, query, selects):  # Select with parameter
        c = self.cursor
        c.execute(query, selects)
        return c.fetchone()
    def select_para_all(self, query, selects):  # Select with parameter
        c = self.cursor
        c.execute(query, selects)
        return c.fetchall()        

