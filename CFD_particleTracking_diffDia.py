# -*- coding: utf-8 -*-
"""
Created on Tue Mar 18 17:29:41 2014

@author: Xiaoning_Shen
"""

import Tkinter,tkFileDialog
import os
rt = Tkinter.Tk()
rt.withdraw()
fileName = tkFileDialog.askopenfilename(parent=rt,title='Choose a file')
fileName_wPath, fileExtension = os.path.splitext(fileName)
outName = ''.join([fileName_wPath, '_Results.xls'])

collectionStatus = ['Incomplete', 'Trapped', 'Escaped', 'Net']
tagName = 'Mass Transfer Summary'

#------------------------------------
def resWrite(fileName, collectionStatus, dataArray, groupInd, groupLabel):
    # This function write the data into an Excel file format
    from xlwt import Workbook, easyxf     
        
    book = Workbook()
    sheet1 = book.add_sheet('Results')
    
    sizeCnt = dataArray.shape[0]
    ngroup = len(groupInd)
    groupInd.append(sizeCnt)
    
    if ngroup != len(groupLabel):
        print 'The length of group label != group number'

    # Format the index label
    i = 0
    for j in xrange(ngroup):
        sheet1.write_merge(i, i, j*5+1, j*5+4, groupLabel[j])
    
    i = 1 
    for j in xrange(ngroup):
        sheet1.row(i).write(j*5+0, '')
        for p in xrange(4):
            sheet1.row(i).write(j*5+p+1, collectionStatus[p])
    # Writing data from line 3
    offset = 2
    for i in range(ngroup):
        for j in range(offset, groupInd[i+1]+offset-groupInd[i]):
            for p in range(4):
                dataInd = j - offset + groupInd[i]
                sheet1.row(j).write(i*5+p+1, dataArray[dataInd, p], easyxf('font: name Arial;', num_format_str='0.00E+00'))

    # Writing total summary
    # ****************'Incomplete'*******************
    offset_total = groupInd[1]-groupInd[0]+5
    sheet1.row(offset_total).write(0, 'Incomplete')
    
    i = offset_total + 1
    for j in xrange(ngroup):
        sheet1.row(i).write(j+1, groupLabel[j])
    offset = offset_total + 2    
    for i in range(ngroup):
        for j in range(offset, groupInd[i+1]+offset-groupInd[i]):
            dataInd = j - offset + groupInd[i]
            sheet1.row(j).write(i+1, dataArray[dataInd, 0], easyxf('font: name Arial;', num_format_str='0.00E+00'))    

    # ****************'Trapped'*******************
    offset_total = offset_total+groupInd[1]-groupInd[0]+5
    sheet1.row(offset_total).write(0, 'Trapped')
    
    i = offset_total + 1
    for j in xrange(ngroup):
        sheet1.row(i).write(j+1, groupLabel[j])
        
    offset = offset_total + 2    
    for i in range(ngroup):
        for j in range(offset, groupInd[i+1]+offset-groupInd[i]):
            dataInd = j - offset + groupInd[i]
            sheet1.row(j).write(i+1, dataArray[dataInd, 1], easyxf('font: name Arial;', num_format_str='0.00E+00'))    
        
    # ****************'Escaped'*******************
    offset_total = offset_total+groupInd[1]-groupInd[0]+5
    sheet1.row(offset_total).write(0, 'Escaped')
    
    i = offset_total + 1
    for j in xrange(ngroup):
        sheet1.row(i).write(j+1, groupLabel[j])
        
    offset = offset_total + 2    
    for i in range(ngroup):
        for j in range(offset, groupInd[i+1]+offset-groupInd[i]):
            dataInd = j - offset + groupInd[i]
            sheet1.row(j).write(i+1, dataArray[dataInd, 2], easyxf('font: name Arial;', num_format_str='0.00E+00')) 
            
    book.save(fileName)
#--------------------------------------


#-------------------------------------
def resRead(fileName, sheetIndex, collectionStatus, tagName):
    # This function read in the data from Excel format and extract results into a numpy array
    import numpy as np
    import re
    from xlrd import open_workbook
    
    book = open_workbook(fileName)
    sheet = book.sheet_by_index(sheetIndex)
    
    dataArray = np.zeros((200, 4))   # [Incomplete, Trapped, Escaped, Net]
    sizeCnt = 0
    groupInd = []
    groupLabel = []                  # group labels
    sepChar = '----'
    # regular searching pattern
    pat = r'\s+(\w+).(-.\w+\s\d+)?\s+(\d+\.\d+[e|E][+|-]\d+)\s+(\d+\.\d+[e|E][+|-]\d+)\s+(\d+\.\d+[e|E][+|-]\d+)' 
        
    groupInd.append(0);
    for j in xrange(sheet.ncols):
        if groupInd[-1] != sizeCnt:
            groupInd.append(sizeCnt)
        print groupInd
        
        i = 0
        if str(sheet.cell_value(i, j)):
            groupLabel.append(str(sheet.cell_value(i, j)))
        print i, j, groupLabel
        
        i = i+1
        while i < sheet.nrows:
            if tagName in sheet.cell_value(i, j):
            # find the 'mass transfer summary' name tag
                k = 0                       # advancing index 
                SingleLineMarker = 0        # Possible single results
                # Read each line of the summary results
                # '- - - -' is the separation line
                while not (sepChar in sheet.cell_value(i+5+k,j)):  
                # still in the detail report
                    match = re.search(pat, sheet.cell_value(i+5+k,j))  
                    # extract the mass fraction
                    if match:  # regular expression matching successful
                        if match.group(1) in collectionStatus:  
                            index = collectionStatus.index(match.group(1))
                            dataArray[sizeCnt, index] = float(match.group(3))
                        else:
                            print 'The initial items are unexpected.'
                            print 'The matched string is: %s' %match.group()
                            print 'The original string is: %s' %sheet.cell_value(i+5+k,j)                       
                    else: # match failed indict the summary has only a single line
                        #print 'Regular expression confronts unexpected patterns. (Reading results)'
                        #print 'The original string is: %s' %sheet.cell_value(i+5+k,j)                       
                        SingleLineMarker = 1     
                        break   # indict "Trapped" is complete
                    k = k + 1
                
                # - - - - 
                # When '- - - -' is met, move the next line
                if SingleLineMarker < 1:
                    k = k+1
                    match = re.search(pat, sheet.cell_value(i+5+k,j))  
                    # extract the mass fraction
                    if match:  # regular expression matching successful
                        if match.group(1) in collectionStatus:  
                            index = collectionStatus.index(match.group(1))
                            if index == 3:
                                dataArray[sizeCnt, index] = float(match.group(3))
                            else:
                                print 'Index is %d' % index
                                print 'The line should start with *Net*'
                                print 'The original string is: %s' %sheet.cell_value(i+5+k,j)                       
                        else:
                            print 'Regular expression confronted unexpected pattern (Reading Net)'
                            print 'The matched string is: %s' %match.group()
                            print 'The original string is: %s' %sheet.cell_value(i+5+k,j)                       
                    else:
                        print 'Regular expression mis-match'
                        print 'The original string is: %s' %sheet.cell_value(i+5+k,j) 
                else:
                    # if SingLeLineMarker on, copy the previous line
                    dataArray[sizeCnt, 3] =  dataArray[sizeCnt, index]
                        
                sizeCnt=sizeCnt+1   # 'Mass Transfer Summary' is found 
                i = i+k+5           # jump the next block
            else:
                i = i+1
                   
    return (dataArray[0:sizeCnt,:], groupInd, groupLabel)
#-------------------------------------

dataArray, groupInd, groupLabel = resRead(fileName, 0, collectionStatus, tagName)
print dataArray.shape
resWrite(outName, collectionStatus, dataArray, groupInd, groupLabel)
print dataArray


