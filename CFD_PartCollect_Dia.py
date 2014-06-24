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
        for j in range(1, len(dataTotal)): # len(dataTotal) is the number of column.
            dataTotal[0][i]=pd.concat([dataTotal[0][i],dataTotal[j][i]], axis=1, join='outer')
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
    Output: dataTotal       -- N x 4 list of 1 x M pandas dataFrame 
                               M: row num, indicting the diameter
                               N: col num, indicting the set parameter
                               4: four types of the collection, i.e. 
                               "Escaped", "Trapped", "Incomplete", "Net"
    '''
    import pandas as pd
    from xlrd import open_workbook

    # global data buffer for processing
    rowName = []
    colName = []
   
    sheetIndex = 0
    book = open_workbook(fileName)
    sheet = book.sheet_by_index(sheetIndex)
   
    dataTotal = []          # store the data of the entire sheet
    
    for colNum in range(sheet.ncols):
        dataGroup = []      # store the data of each column
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
                    
                    # extract the data and row name                       
                    dataSeg = chunk_process(chunk)
                    
                    # initiate the entire data array if not exist yet.                        
                    if len(dataGroup) == 0:
                        for i in range(4): # "Escaped", "Trapped", "Incomplete", "Net"
                            dataGroup.append(pd.DataFrame(dataSeg[i], index=rowName, \
                                                columns=[sheet.cell_value(0, colNum)]))
                    # if the final result array exists, merge with the current new data.
                    else:
                        for i in range(4):
                            dataTemp.append(pd.DataFrame(dataSeg[i], index=rowName, \
                                                columns=[sheet.cell_value(0, colNum)]))
                            dataGroup[i] = pd.concat([dataGroup[i], dataTemp[i]], axis=0, join='inner')
                         # clear dataTemp for next time.
                        dataTemp = []                         
                    # jump over the analyzed chunk.
                    ind += counter
            # remove the duplicate indices
            for i in range(4):
                df = dataGroup[i]
                df["index"] = df.index
                df.drop_duplicates(cols='index', take_last=True, inplace=True)
                del df["index"]
                dataGroup[i] = df.sort()

            dataTotal.append(dataGroup) 
            # ALl chunked at each col is loaded to the memory   
    return dataTotal

def chunk_process(chunk):
    '''
    Extract the collection data of each chunk from Excel 
    '''
    # labeling the final results of the particle tracking.
    res_tag = re.compile('Mass Transfer Summary')    
    # characters for separating counting lines
    collectionStatus = ['Incomplete', 'Trapped', 'Escaped', 'Net']
    lineOffset = 3
    # regular expression for " Trapped - Zone 22       1.042e-04  1.042e-04  0.000e+00"
    pat = r'\s+(\w+).(-.\w+\s\d+)?\s+(\d+\.\d+[e|E][+|-]\d+)\s+(\d+\.\d+[e|E][+|-]\d+)\s+(\d+\.\d+[e|E][+|-]\d+)' 

    # buffer results for each segment
    dataArray = np.zeros(4)
    
    # extract 'injection-75-141' style expression    
    # looping through the given column
    for i in range(len(chunk)):
        # This is the injection results section
        # starting here "Trapped", "Escaped", "Incomplete" & "Net" are analyzed
        # separately
        if (res_tag.search(chunk[i]) != None):
        # find the 'mass transfer summary' name tag
            k = 0                       # advancing index 
            ResLineNum = 0              # Possible single results
            # Read each line of the summary results
            # '- - - -' is the separation line
            # The condition is a) the search less than the total line number
            #                  b)'----' is not in the line
            # Bascially, search between '----' and '-----'
            # '5' is the preset offset
            while ((i+k+lineOffset) < len(chunk)):  
                # Search for reports on "Trapped", "Escaped", "Incomplete" & "Net" 
                match = re.search(pat, chunk[i+k+lineOffset])
                # extract the mass fraction
                if match:  # regular expression matching successful
                    if match.group(1) in collectionStatus:  
                        index = collectionStatus.index(match.group(1))
                        dataArray[index] = float(match.group(3))
                        # If there is only one line, then there is no 'Net' line below                    
                        ResLineNum += 1
                    else:
                        print 'The initial items are unexpected. Change the lineOffset value.'
                        print 'The matched string is: %s' %match.group()
                        print 'The original string is: %s' %chunk[i+k+lineOffset]
                k = k + 1

            if ResLineNum < 2:
                dataArray[3] =  dataArray[index]

            # When '- - - -' is met, move the next line
            # If only one line is present, then no need to analyzing the following line.
#            if ResLineNum > 1:
#                k = k + 1     # jump over '---' line
#                match = re.search(pat,chunk[i+k+lineOffset])  
#                # extract the mass fraction
#                if match:  # regular expression matching successful
#                    if match.group(1) in collectionStatus:  
#                        index = collectionStatus.index(match.group(1))
#                        if index == 3:          # This line has to be 'Net'
#                            dataArray[index] = float(match.group(3))
#                        else:
#                            print 'Index is \" %s \"' % collectionStatus[index]
#                            print 'The line should start with *Net*'
#                            print 'The original string is: %s' %chunk[i+lineOffset+k]
#                    else:
#                        print 'Regular expression finds items out of four choices (Reading Net)'
#                        print 'The matched string is: %s' %match.group()
#                        print 'The original string is: %s' %chunk[i+lineOffset+k]                      
#                        print 'The original string is: %s' %chunk[i+lineOffset+k]                 
#                else:
#                    print 'Regular expression mis-match. The *Net* line is not formatted correctly.'
#                    print 'The original string is: %s' %chunk[i+lineOffset+k]
#            else:
#                # if SingLeLineMarker on, copy the previous line
#                dataArray[3] =  dataArray[index]

    return dataArray

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
