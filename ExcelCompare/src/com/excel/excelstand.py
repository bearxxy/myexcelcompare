'''
Created on Jan 26, 2015

@author: xionxiao
@attention: # Open an existing excel file or create one.
            # Get cell value by given row and column number.
            # Set value to a certain cell.
            # Get row count by given column number.
            # Get column count by given row number. 
'''
import os
import sys

# from django.utils.termcolors import background
# from wx.tools.Editra.src.plugdlg import MODE_CONFIG
import xlrd
from xlutils.copy import copy
import xlwt


class ExcelApp:
    '''
    classdocs
    '''
    dataDir = ""
    fileName = ""
    sheetsName = ["Sheet1"]
  
    def __init__(self, fileName, sheetsName):
        '''
        Constructor
        '''
        self.fileName = fileName
        self.sheetsName = sheetsName

        baseDir = sys.path[0]
        dataDir = ""
        
        if os.path.isdir(baseDir):
            dataDir =  baseDir
        elif os.path.isfile(baseDir):
            dataDir =  os.path.dirname(baseDir)
            
        if not os.path.isfile(fileName):            
            reportBook = xlwt.Workbook()
            for sht in self.sheetsName:
                reportBook.add_sheet(sht)
#             self.fileName = "".join([dataDir.replace("\\", "/"), "/", self.fileName])
            print self.fileName
            reportBook.save(self.fileName)

    #===========================================================================
    # openExcel
    #===========================================================================
    def openExcel(self, mode):
        if mode == 'w':
            return copy(xlrd.open_workbook(self.fileName))
        elif mode == 'r': 
            return xlrd.open_workbook(self.fileName)

    
    #===========================================================================
    # getSheetByIndex
    #===========================================================================
    def getSheetByIndex(self,book, sheetIndex):
        sheet = book.get_sheet(sheetIndex)
        return sheet
    
    
    #===========================================================================
    # getSheetByName
    #===========================================================================
    def getSheetByName(self, book, sheetName):
        sheet = book.sheet_by_name(sheetName)
        return sheet
    
    #===========================================================================
    # saveExcel
    #===========================================================================
    def saveExcel(self, book):
        book.save(self.fileName)
#         book.close()

    #===========================================================================
    # getSheetsName
    #===========================================================================
    def getSheetsName(self):
        sheetlist = []
        book = xlrd.open_workbook(self.fileName)
        for sheet in book.sheets():
            sheetlist.append(sheet.name)

        return sheetlist
        
    #===========================================================================
    # setCellData
    #===========================================================================
    def setCellData(self, sheetObj, row, col, cellValue, highlight = False, backcolor= "white", fontcolor = "black", bold = "off"):

        if highlight == False:
            sheetObj.write(row, col, cellValue)
        else:
            style = xlwt.easyxf('pattern: pattern solid, fore-color '+ backcolor+'; font:color-index '+ fontcolor +',bold '+ bold+',name HP Simplified')
            sheetObj.write(row, col, cellValue, style)

    def setColWidth(self,sheetObj, colNo, widthSize):
        sheetObj.col(colNo).width = widthSize
    
    #===========================================================================
    # getRowData
    # Mode: read
    #===========================================================================
    def getRowData(self, sheetObj, row):
        return sheetObj.row_values(row)
    
    #===========================================================================
    # getColData
    # Mode: read
    #===========================================================================
    def getColData(self, sheetObj, col):
        return sheetObj.col_values(col)
        
    #===========================================================================
    # getCellData
    #===========================================================================
    def getCellData(self, sheetObj, row, col):

        cellData = ""
        try:
#             data = xlrd.open_workbook(self.fileName)
            cellData = sheetObj.cell(row, col).value
        finally:
            return cellData
                
    #===========================================================================
    # getRowCount
    #===========================================================================
    def getRowCount(self,sheetName, col):
            return xlrd.open_workbook(self.fileName).sheet_by_name(sheetName).nrows
        
    #===========================================================================
    # getColCount
    #===========================================================================
    def getColCount(self, sheetName, row):
        return xlrd.open_workbook(self.fileName).sheet_by_name(sheetName).ncols


if __name__=="__main__":
    
#     excel2 = ExcelApp("C:/Users/xionxiao/Desktop/222.xls",["Compare Result"])
#     book = excel2.openExcel('r')
#     print excel2.getCellData(excel2.getSheetByName(book,"Sheet 1"), 0, 0)
#     excel = ExcelApp("C:/Users/xionxiao/Desktop/222.xls",["Compare Result"])
#     aBook = excel.openExcel('w')
#     aSheet = excel.getSheetByIndex(aBook, 0)
#     print "***"+excel.getCellData(aSheet, 0, 0)
#     excel.setCellData(aSheet,2, 4, "sssss")
#     excel.saveExcel(aBook)
    a = ["4","3","5"]
    a.sort()
    print a
    
    d = {4:['5','1','3'],2:['3','5','2']}
    print sorted(d)
    
        
#     excel.setCellData(0, 1, 1, "test", True)
        