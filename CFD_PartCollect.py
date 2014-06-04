# -*- coding: utf-8 -*-
"""
Created on Fri Apr 25 11:02:08 2014

Script that handles CFD output files

Based on Xiaoning's Code

Not optimized: basic info should be grabbed once and for all

@author: Rongxiao Zhang
"""

import Tkinter,tkFileDialog
import os
import re

def readExcel(fileName, tag):
    '''
    reads and extracts info from the given excel file
    Input:  fileName -- file name including the directory path
            tag -- Identifier to the next reading block
    Output: dataCollection  -- M x N x 4 
                               M: row num, indicting the diameter
                               N: col num, indicting the set parameter
                               4: four types of the collection, i.e. 
                               "Escaped", "Trapped", "Incomplete", "Net"
            rowName         -- 1 x M list of the diameter
            colName         -- 1 x N list of the set parameter
    '''
    import numpy as np
    from xlrd import open_workbook

    # global data buffer for processing
    maxRow = 10000l
    maxCol = 10000l
    dataCollection = np.empty([maxRow, maxCol, 4])    
    rowName = []
    colName = []
   
    sheetIndex = 0
    book = open_workbook(fileName)
    sheet = book.sheet_by_index(sheetIndex)
    host = {}
    basicInfo = []
    

    
    # for each column, read in the entire section into memory for processing    
    for colNum in range(sheet.ncols):  # loop through columns
        col_counter = 0l
        if (sheet.cell(0, colNum) != ''):
            # if the header string exist, continue reading the column
            # otherwise jump the next column
            col_values = sheet.col_values(colNum) # grab values
            # Grab the data marked between tags
            for ind, x in enumerate(col_values): 
                if (tag.search(x) != None): 
                    # find the next chunk of data 
                    chunk = []
                    counter = 1
                    
                    # above the sheet bottom & not at the next block
                    while((ind + counter) < sheet.nrows and \
                    (tag.search(col_values[ind + counter]) == None):                        
                        chunk.append(col_values[ind + counter])
                        counter += 1
                        basicInfo = chunk_process(chunk, host)
                    
                    # load the data into the dataCollection
                    dataCollection[]
        
                    i += counter

            col_counter = col_counter+1
        
    return host, basicInfo

def chunk_process(chunk, host):
    '''
    Extract the collection data of each chunk from Excel 
    '''
    import numpy as np
    # scienfic notation. ?: indicts non-capturing group
    sn_tag = re.compile('[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?')
    # injection information in the style of 'injection-75-142.9  198'
    r_tag = re.compile('injection-(?:[-+]?[0-9]*\.?[0-9]+)-([-+]?[0-9]*\.?[0-9]+)')
    # trapped_tag = re.compile('Trapped - Zone')
    escaped_tag = re.compile('Escaped - Zone')
    # labeling the final results of the particle tracking.
    res_tag = 'Mass Transfer Summary'    
    sepChar = '----'
    collectionStatus = ['Incomplete', 'Trapped', 'Escaped', 'Net']
    # regular expression for " Trapped - Zone 22       1.042e-04  1.042e-04  0.000e+00"
    pat = r'\s+(\w+).(-.\w+\s\d+)?\s+(\d+\.\d+[e|E][+|-]\d+)\s+(\d+\.\d+[e|E][+|-]\d+)\s+(\d+\.\d+[e|E][+|-]\d+)' 

    dataArray = np.zeros((500, 4))
    diaList = [] # diameter and angle
#   trapped_temp = {}    
    escaped_temp = {}
    
    # extract 'injection-75-141' style expression    
    # looping through the given column
    for i in range(np.size(chunk)):
        # find 'Escaped - Zone' line -- input information
        if (escaped_tag.search(chunk[i]) != None):
            # find injection information 'injection-75-142.9  198'
            if(r_tag.search(chunk[i]) != None):
                # This is the injection info section
                # extract the numbers of the 'injection...' string
                r_list  = {np.float(x) for x in r_tag.findall(chunk[i])[0]}
                dia     = r_list[0]
                # forming the row name
                diaList.append(dia)
        # find '
        # This is the injection results section
        # starting here "Trapped", "Escaped", "Incomplete" & "Net" are analyzed
        # separately
        if (res_tag.search(chunk[i]) != None):
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

    # This is the injection results section
    # starting here "Trapped", "Escaped", "Incomplete" & "Net" are analyzed
    # separately

    
def writeExcel(fileName, host, basicInfo):
    import numpy as np
    from xlwt import Workbook, Formula, easyxf
    from xlrd import cellname
    
    book = Workbook()
    sheet1 = book.add_sheet('Results') 
    
    # Basic Info
    sheet1.row(0).write(0, 'Diameter', easyxf('font: name Arial;'))
    sheet1.row(0).write(1, basicInfo[0], easyxf('font: name Arial;'))
    sheet1.row(0).write(2, 'Release Angle', easyxf('font: name Arial;'))
    sheet1.row(0).write(3, basicInfo[1], easyxf('font: name Arial;'))
    
    keyList = []
    groupList = []

    # sort the keylist
    for key, value in host.iteritems():
        keyList.append(key)
        tempGroupList = []
        for key2, value2 in value.iteritems():
            tempGroupList.append(key2)
        if np.size(tempGroupList) > np.size(groupList):
            groupList = np.sort(tempGroupList)

    keyList = np.sort(keyList)
    # set column width
    for i in range((np.size(groupList) + 1) * 2):
        sheet1.col(i).width = 4000
    
    for i in range(np.size(keyList)):
        # writing r
        sheet1.row(i+2).write(0, keyList[i], easyxf('font: name Arial;'))
        sheet1.row(i+2).write(3 + np.size(groupList), keyList[i], \
                              easyxf('font: name Arial;'))
        
        sheet1.row(i+2).write(np.size(groupList) + 1,\
                              Formula('SUM(%s:%s)' % (cellname(i+2, 1), \
                              cellname(i+2, np.size(groupList)))), \
                              easyxf('font: name Arial;', num_format_str='0.00E+00'))
                              
        sheet1.row(i+2).write(2 * np.size(groupList) + 4, \
                              Formula('SUM(%s:%s)' % \
                              (cellname(i+2, 4 + np.size(groupList)), \
                              cellname(i+2, 3 + 2 * np.size(groupList)))), \
                              easyxf('font: name Arial;', num_format_str='0.00%'))
                              
        for j in range(np.size(groupList)):
            try:
                sheet1.row(i+2).write(j+1, host[keyList[i]][groupList[j]][0], \
                                      easyxf('font: name Arial;', \
                                      num_format_str='0.00E+00'))
            except KeyError:
                pass
            
    for i in range(np.size(groupList)):
        # writing group names
        sheet1.row(1).write(i + 1, groupList[i], easyxf('font: name Arial;'))
        sheet1.row(1).write(i + 4 + np.size(groupList), groupList[i], \
                            easyxf('font: name Arial;'))
        
        for j in range(np.size(keyList)):
            sheet1.row(j + 2).write(i + 4 + np.size(groupList), \
                                    Formula('%s / %s' % (cellname(j+2, i+1), \
                                    cellname(j+2, np.size(groupList) + 1))), \
                                    easyxf('font: name Arial;', num_format_str='0.00%'))
        
    sheet1.row(1).write(np.size(groupList) + 1, 'Total', \
                        easyxf('font: name Arial;'))
    sheet1.row(1).write(np.size(groupList) * 2 + 4 , 'Total', \
                        easyxf('font: name Arial;'))
    
    book.save(fileName)   

def main():
    rt = Tkinter.Tk()
    rt.withdraw()
    fileName = tkFileDialog.askopenfilename(parent=rt,title='Choose a file')
    fileName_wPath, fileExtension = os.path.splitext(fileName)
    outName = ''.join([fileName_wPath, '_Results.xls'])
    
    tag = re.compile('number tracked')
    host, basicInfo = readExcel(fileName, tag)
    writeExcel(outName, host, basicInfo)

if __name__ == '__main__':
    main()
