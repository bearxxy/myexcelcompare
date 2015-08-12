'''
Created on Jul 20, 2015

@author: xionxiao
'''

import time
from com.excel.excelstand import ExcelApp

def compareFiles(sourceFile, sourceSheet, targetFile, targetSheet, reportName, pkNO):
    srcExcel = ExcelApp(sourceFile,"")
    tgtExcel = ExcelApp(targetFile,"")
    resultExcel = ExcelApp(reportName, ["Compare_Result","Records_in_Source_NOT_in_Target","Records_in_Target_NOT_in_Source"])
    
    srcColCount = srcExcel.getColCount(sourceSheet, 0)
    srcRowCount = srcExcel.getRowCount(sourceSheet, 0)
    tgtRowCount = tgtExcel.getRowCount(targetSheet, 0)
    
    srcbook = srcExcel.openExcel("r")
    tgtbook =  tgtExcel.openExcel("r")
    rstbook = resultExcel.openExcel("w")
    
    srcsheet = srcExcel.getSheetByName(srcbook, sourceSheet)
    tgtsheet = tgtExcel.getSheetByName(tgtbook, targetSheet)
    rstsheet = resultExcel.getSheetByIndex(rstbook, 0)
    missingInSRCSheet = resultExcel.getSheetByIndex(rstbook, 1)
    missingInTGTSheet = resultExcel.getSheetByIndex(rstbook, 2)
    
    srcdic = {}
    tgtdic = {}
        
    k = 0   
    i = 0
    j = 0
    
    
    print time.strftime('%I-%M-%S_%Y-%m-%d',time.localtime(time.time()))
    
    ''' 
    Write missing records to result report
    '''
    while i < srcRowCount:
        srcrowdata = srcExcel.getRowData(srcsheet, i)
        srcdic.setdefault(srcExcel.getCellData(srcsheet, i, pkNO), []).append(srcrowdata)
        i = i + 1
#     print sorted(srcdic.iteritems())
    
    while j < tgtRowCount:
        tgtrowdata = tgtExcel.getRowData(tgtsheet, j)
        tgtdic.setdefault(tgtExcel.getCellData(tgtsheet, j, pkNO), []).append(tgtrowdata)
        j = j + 1
    
    '''
    1. check common records
    2. write common records 
    3. delete common records and leave missing records
    '''   
    i = 0
    j = 0
    r = 5
    commonRecords = 0;
    matchCount = [0]
    differCount = [0]
    for k in range(srcColCount):
        matchCount.append(0)
        differCount.append(0)
        
    sortedSrcDic = sorted(srcdic.iteritems())
    sortedTgtDic = sorted(tgtdic.iteritems())

    while i < len(sortedSrcDic) and j < len(sortedTgtDic):
        
        if sortedSrcDic[i][0] > sortedTgtDic[j][0]:
            j = j + 1
        elif sortedSrcDic[i][0] < sortedTgtDic[j][0]:
            i = i + 1
        else:
            commonRecords = commonRecords + 1
#             resultExcel.setCellData(rstsheet, r, 1, sortedSrcDic[i][0])
#             resultExcel.setCellData(rstsheet, r + 1, 1, sortedTgtDic[j][0])
  
            k = 0
            while k < srcColCount:
                if sortedSrcDic[i][1][0][k] == sortedTgtDic[j][1][0][k]:
                    resultExcel.setCellData(rstsheet, r, k + 1, sortedSrcDic[i][1][0][k])
                    resultExcel.setCellData(rstsheet, r + 1, k + 1, sortedTgtDic[j][1][0][k])
                    matchCount[k] = matchCount[k] +  1
                else:
                    resultExcel.setCellData(rstsheet, r, k + 1, sortedSrcDic[i][1][0][k], True,"light-orange", "black")                
                    resultExcel.setCellData(rstsheet, r + 1, k + 1, sortedTgtDic[j][1][0][k], True,"orange", "black")
                    differCount[k] = differCount[k] + 1
                k = k + 1
            r = r + 3
            del sortedSrcDic[i]
            del sortedTgtDic[j]
    
    ''' 
    Write missing records to result report
    '''
    i = 0
    while i < len(sortedSrcDic):
        j = 0
        while j < srcColCount:
            resultExcel.setCellData(missingInSRCSheet, i, j, sortedSrcDic[i][1][0][j])
            j = j + 1
        i = i + 1
        
    i = 0
    while i < len(sortedTgtDic):
        j = 0
        while j < srcColCount:
            resultExcel.setCellData(missingInTGTSheet, i, j, sortedTgtDic[i][1][0][j])
            j = j + 1
        i = i + 1
    
    ''' 
    Summary
    '''  
    resultExcel.setCellData(rstsheet, 0, 0, "Good Records on Field",True,"white","light-blue", "on")
    resultExcel.setCellData(rstsheet, 1, 0, "Good Records Rate on Field",True,"white","light-blue", "on")
    resultExcel.setCellData(rstsheet, 2, 0, "Bad Records on Field",True,"white","light-blue", "on")
    resultExcel.setCellData(rstsheet, 3, 0, "Bad Records Rate on Field",True,"white","light-blue", "on")
    resultExcel.setColWidth(rstsheet, 0, 6666)
    
    for k in range(srcColCount):        
        if not differCount[k] == 0:
            color =  "red"
        else:
            color = "green"

        resultExcel.setCellData(rstsheet, 0, k + 1, matchCount[k], True, "green","white", "on")
        resultExcel.setCellData(rstsheet, 1, k + 1, "".join([str(matchCount[k]*100/commonRecords),"%"]),True,"green","white", "on")      
        resultExcel.setCellData(rstsheet, 2, k + 1, differCount[k], True, color,"white", "on")
        resultExcel.setCellData(rstsheet, 3, k + 1, "".join([str(differCount[k]*100/commonRecords),"%"]), True,color,"white", "on")
    
    ''' 
    Save report
    '''
    resultExcel.saveExcel(rstbook)

    print time.strftime('%I-%M-%S_%Y-%m-%d',time.localtime(time.time()))
