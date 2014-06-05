# -*- coding: utf-8 -*-
"""
Created on Wed Jun 04 14:59:16 2014

@author: Rongxiao.Zhang
"""

import Tkinter,tkFileDialog
import os
import re
import pandas as pd
import numpy as np


def merge_dataFrame(dataTotal):
    for i in range(4):
        for j in range(1, len(dataTotal)):
            dataTotal[0][i] = pd.concat([dataTotal[0][i], dataTotal[j][i]], axis=1, join='inner')
    return dataTotal[0][:]

def writeExcel(file_name, df_list):
    w = pd.ExcelWriter(file_name)
    num_groups = 4 # input is MxNx4
    names = ['Incomplete', 'Trapped', 'Escaped', 'Net']
    start_row = 0
    
    for item, i in zip(df_list, range(num_groups)):
        size_x = item.shape[0]
        item.to_excel(w, startrow = start_row, startcol = 0, index_label=names[i])
        start_row += (size_x + 3)

    w.save()

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
    import pandas as pd
    from xlrd import open_workbook

    # global data buffer for processing
    rowName = []
    colName = []
   
    sheetIndex = 0
    book = open_workbook(fileName)
    sheet = book.sheet_by_index(sheetIndex)
   
    dataTotal = []
    for colNum in range(sheet.ncols):
        dataGroup = []
        dataTemp = []
            
        if (sheet.cell_value(0, colNum) != ''):
            # if the header string exist, continue reading the column
            # otherwise jump the next column
    
            colName.append(sheet.cell_value(0, colNum))
            col_values = sheet.col_values(colNum) # grab values
            # Grab the data marked between tags
            for ind, x in enumerate(col_values): 
                if (tag.search(x) != None): 
                    # find the next chunk of data 
                    chunk = []
                    counter = 1
                    
                    # above the sheet bottom & not at the next block
                    while((ind + counter) < sheet.nrows and \
                    (tag.search(col_values[ind + counter]) == None)):                        
                        chunk.append(col_values[ind + counter])
                        counter += 1
                        
                    dataSeg, rowName = chunk_process(chunk)
                        
                     
                    if dataGroup == []:
                        for i in range(4):
                            dataGroup.append(pd.DataFrame(dataSeg[:,i], index=rowName, \
                                                columns=[sheet.cell_value(0, colNum)]))
                    else:
                        for i in range(4):
                            dataTemp.append(pd.DataFrame(dataSeg[:,i], index=rowName, \
                                                columns=[sheet.cell_value(0, colNum)]))     
                            dataGroup[i] = pd.concat([dataGroup[i], dataTemp[i]], axis=0, join='inner')
                        dataTemp = []                         
                                    
                        ind += counter
    
            dataTotal.append(dataGroup) 
                        # ALl chunked at each col is loaded to the memory   
    return dataTotal

def chunk_process(chunk):
    '''
    Extract the collection data of each chunk from Excel 
    '''
    # injection information in the style of 'injection-75-142.9  198'
    r_tag = re.compile('injection-(?:[-+]?[0-9]*\.?[0-9]+)-([-+]?[0-9]*\.?[0-9]+)')
    # trapped_tag = re.compile('Trapped - Zone')
    escaped_tag = re.compile('Escaped - Zone')
    trapped_tag = re.compile('Trapped - Zone')
    # labeling the final results of the particle tracking.
    res_tag = re.compile('Mass Transfer Summary')    
    # characters for separating counting lines
    sepChar = '----'
    collectionStatus = ['Incomplete', 'Trapped', 'Escaped', 'Net']
    sizeCnt = 0
    # regular expression for " Trapped - Zone 22       1.042e-04  1.042e-04  0.000e+00"
    pat = r'\s+(\w+).(-.\w+\s\d+)?\s+(\d+\.\d+[e|E][+|-]\d+)\s+(\d+\.\d+[e|E][+|-]\d+)\s+(\d+\.\d+[e|E][+|-]\d+)' 

    # buffer results for each segment
    dataArray = np.zeros((500, 4))
    diaList = []    # diameter list
    
    # extract 'injection-75-141' style expression    
    # looping through the given column
    for i in range(np.size(chunk)):
        # find 'Escaped - Zone' line -- input information
        if (escaped_tag.search(chunk[i]) != None or trapped_tag.search(chunk[i])):
            # find injection information 'injection-75-142.9  198'
            if(r_tag.search(chunk[i]) != None):

                # This is the injection info section
                # extract the numbers of the 'injection...' string
                dia  = r_tag.findall(chunk[i])[0]
                # forming the row name
                diaList.append(dia)

        # This is the injection results section
        # starting here "Trapped", "Escaped", "Incomplete" & "Net" are analyzed
        # separately
        if (res_tag.search(chunk[i]) != None):
        # find the 'mass transfer summary' name tagn
            k = 0                       # advancing index 
            SingleLineMarker = 0        # Possible single results
            # Read each line of the summary results
            # '- - - -' is the separation line
            while ((i+k+5) < 12 and (sepChar not in chunk[i+k+5])):  
            # still in the detail report
                match = re.search(pat, chunk[i+k+5])
                # extract the mass fraction
                if match:  # regular expression matching successful
                    if match.group(1) in collectionStatus:  
                        index = collectionStatus.index(match.group(1))
                        dataArray[sizeCnt, index] = float(match.group(3))
                        SingleLineMarker += 1
                    else:
                        print 'The initial items are unexpected.'
                        print 'The matched string is: %s' %match.group()
                        print 'The original string is: %s' %chunk[i+k+5]                       

                k = k + 1
            
            # - - - - 
            # When '- - - -' is met, move the next line
            if SingleLineMarker < 1:
                k = k+1
                match = re.search(pat,chunk[i+5+k])  
                # extract the mass fraction
                if match:  # regular expression matching successful
                    if match.group(1) in collectionStatus:  
                        index = collectionStatus.index(match.group(1))
                        if index == 3:
                            dataArray[sizeCnt, index] = float(match.group(3))
                        else:
                            print 'Index is %d' % index
                            print 'The line should start with *Net*'
                            print 'The original string is: %s' %chunk[i+5+k]
                    else:
                        print 'Regular expression confronted unexpected pattern (Reading Net)'
                        print 'The matched string is: %s' %match.group()
                        print 'The original string is: %s' %chunk[i+5+k]                      
                        print 'The original string is: %s' %chunk[i+5+k]                 
                else:
                    print 'Regular expression mis-match'
                    print 'The original string is: %s' %chunk[i+5+k]
            else:
                # if SingLeLineMarker on, copy the previous line
                dataArray[sizeCnt, 3] =  dataArray[sizeCnt, index]
                    
            sizeCnt=sizeCnt+1   # 'Mass Transfer Summary' is found 
            i = i+1             # jump the next block
        else:
            i = i+1

    return dataArray[0:sizeCnt,:], [dia]

def main():
    rt = Tkinter.Tk()
    rt.withdraw()
    fileName = tkFileDialog.askopenfilename(parent=rt,title='Choose a file')
    fileName_wPath, fileExtension = os.path.splitext(fileName)
    outName = ''.join([fileName_wPath, '_Results.xls'])
    
    tag = re.compile('number tracked')
    dataGroup = readExcel(fileName, tag)  
    dataGroup = merge_dataFrame(dataGroup)
    writeExcel(outName, dataGroup)

if __name__ == '__main__':
    main()
