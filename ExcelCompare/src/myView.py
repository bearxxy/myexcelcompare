'''
Created on Jul 20, 2015

@author: xionxiao
'''

import os
import sys
import time

from PyQt4 import QtCore
import PyQt4
from PyQt4.QtGui import QMainWindow, QApplication, QLabel, QGridLayout, QWidget, \
    QLineEdit, QPushButton, QComboBox, QFileDialog, QIcon, QFont
import win32api

from com import compare
from com.compare.excelCompare import compareFiles
from com.excel.excelstand import ExcelApp


from PyQt4.QtCore import *;

sourceFilePath = ""

class MyView(QMainWindow):
    '''
    MyView: UI as communication with users
    '''
    reportName = ""

    def __init__(self, parent=None):
        '''
        Constructor
        '''
        super(MyView, self).__init__(parent)
       
        ft = QFont("HP Simplified", 9, QFont.Normal)
        self.sourceFileLbl = QLabel("Source File Path:")
        self.sourceSheetLbl = QLabel("Source Sheet:")
        self.targetFileLbl = QLabel("Target File Path:")
        self.targetSheetLbl = QLabel("Target Sheet:")
        self.primaryKeyLbl = QLabel("Primary Key Column:")
#         
#         self.sourceFileLbl.setFont(ft)
#         self.sourceSheetLbl.setFont(ft)
#         self.targetFileLbl.setFont(ft)
#         self.targetSheetLbl.setFont(ft)
#         self.primaryKeyLbl.setFont(ft)
        
        self.sourcePathEdit = QLineEdit()
        self.targetPathEdit = QLineEdit()
        self.primaryKeyEdit = QLineEdit("1")

        
        self.soureBtn = QPushButton("Select Source")
        self.targetBtn = QPushButton("Select Target")
        self.CompareBtn = QPushButton("Compare")
        self.quitBtn = QPushButton("Quit")
        self.clearBtn = QPushButton("Clear")
        self.checkReportBtn = QPushButton("Open Report")
        self.checkReportBtn.setEnabled(False)
        self.soureBtn.setToolTip("Click to select the source file to be compared")
        self.targetBtn.setToolTip("Click to select the target file to be compared")
        self.CompareBtn.setToolTip("Click to start compare")
        self.clearBtn.setToolTip("Clear all text")
        self.quitBtn.setToolTip("Exit from application")
        
        self.sourceSheetBox = QComboBox()
        self.targetSheetBox = QComboBox()
#         self.sourceSheetList = QStringList()
#         self.targetSheetList = QStringList()
        
        self.connect(self.soureBtn, SIGNAL("clicked()"),self.openFile)
        self.connect(self.targetBtn, SIGNAL("clicked()"),self.openFile)
        self.connect(self.CompareBtn, SIGNAL("clicked()"),self.compare)
        self.connect(self.quitBtn, SIGNAL("clicked()"),self.exitApp)
        self.connect(self.clearBtn, SIGNAL("clicked()"),self.clearApp)
        self.connect(self.checkReportBtn, SIGNAL("clicked()"),self.openReport)
        
        layout = QGridLayout()
        layout.addWidget(self.sourceFileLbl, 0, 0)
        layout.addWidget(self.sourceSheetLbl, 1, 0)
        layout.addWidget(self.targetFileLbl, 2, 0)
        layout.addWidget(self.targetSheetLbl, 3, 0)
        layout.addWidget(self.primaryKeyLbl, 4, 0)
        
        layout.addWidget(self.sourcePathEdit, 0, 1)
        layout.addWidget(self.targetPathEdit, 2, 1)
        layout.addWidget(self.primaryKeyEdit, 4, 1)
        
        layout.addWidget(self.soureBtn, 0, 2)
        layout.addWidget(self.targetBtn, 2, 2)
        layout.addWidget(self.CompareBtn, 5, 1)
        layout.addWidget(self.checkReportBtn, 6, 1)        
        layout.addWidget(self.clearBtn, 5, 2)
        layout.addWidget(self.quitBtn, 6, 2)
        
        layout.addWidget(self.sourceSheetBox, 1, 1)
        layout.addWidget(self.targetSheetBox, 3, 1)
        layout.setSpacing(5)
        layout.setMargin(30)
        wgt = QWidget()       
        wgt.setLayout(layout)
        self.setCentralWidget(wgt)
        self.setFont(ft)
        self.resize(500,300)
        self.setWindowTitle("ExcelCompareTool")
        self.setWindowIcon(QIcon("C:/Users/xionxiao/Pictures/ico/katomic.ico"))
#         self.setToolTip(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
#         self.setWindowFlags(QtCore.Qt.FramelessWindowHint)

    def openReport(self):
        os.startfile(self.reportName)

    def clearApp(self):
        self.sourcePathEdit.clear()
        self.targetPathEdit.clear()
        self.sourceSheetBox.clear()
        self.targetSheetBox.clear()
        self.checkReportBtn.setEnabled(False)

    def exitApp(self):   
        if win32api.MessageBox(0,"You will exit from this application, are you sure?","",1) == 1:        
            sys.exit();

    def openFile(self):  
        s = QFileDialog.getOpenFileName(self,"Open file dialog","/","Excel files(*.xls)")
        if not s == "":
            if self.sender().text() == "Select Source":
                self.sourcePathEdit.setText(str(s))
                self.sourceSheetBox.clear()             
                sourceExcel = ExcelApp(s,"")
                self.sourceSheetBox.addItems(sourceExcel.getSheetsName())
                
            elif self.sender().text() == "Select Target":                
                self.targetPathEdit.setText(str(s))
                self.targetSheetBox.clear()
                targetExcel = ExcelApp(s,"")
                self.targetSheetBox.addItems(targetExcel.getSheetsName())

    def compare(self):
        if not self.sourcePathEdit.text() == "" and not self.sourcePathEdit.text() == "":
            self.checkReportBtn.setEnabled(False)
            if not os.path.exists("C:/ExcelCompareReport"):
                os.makedirs("C:/ExcelCompareReport")
            self.reportName = "C:/ExcelCompareReport/Report_"+time.strftime('%I-%M-%S_%Y-%m-%d',time.localtime(time.time()))+".xls"
            compareFiles(self.sourcePathEdit.text(), self.sourceSheetBox.currentText(), self.targetPathEdit.text(), \
                         self.targetSheetBox.currentText(), self.reportName, int(self.primaryKeyEdit.text())-1)
            win32api.MessageBox(0,"Compare report has generated!","",0)
            self.checkReportBtn.setEnabled(True)
        else:
            win32api.MessageBox(0,"Source file or target file shouldn't be blank!","",0)
        
if __name__ == "__main__":
    
    app = QApplication(sys.argv)
    mw = MyView()
    mw.show()
    sys.exit(app.exec_())