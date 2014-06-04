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
    '''
    import numpy as np
    from xlrd import open_workbook
    
    sheetIndex = 0
    book = open_workbook(fileName)
    sheet = book.sheet_by_index(sheetIndex)
    host = {}
    basicInfo = []
    
    for colNum in range(np.size(sheet.row_values(0))):
        # loop through columns. End of file is defined by two adjacent empty columns
        if (sheet.cell(0, colNum) != ''):
            col_values = sheet.col_values(colNum) # grab values
            
            # Grab the data marked between tags
            for i in range(np.size(col_values)):
                if (tag.search(str(col_values[i])) != None):
                    
                    chunk = []
                    counter = 1
                    
                    while((i + counter) < np.size(col_values) and \
                    tag.search(str(col_values[i + counter])) == None):                        
                        chunk.append(col_values[i + counter])
                        counter += 1
                        basicInfo = chunk_process(chunk, host)
                        
                    i += counter
    
    return host, basicInfo

def chunk_process(chunk, host):
    '''
    Grab infomation in the chunks given
    !!!!!
    Doesn't handle trapped particle case yet
    !!!!!
    '''
    import numpy as np
    # scienfic notation. ?: indicts non-capturing group
    sn_tag = re.compile('[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?')
    # injection information in the style of 'injection-75-142.9  198'
    r_tag = re.compile('injection-([-+]?[0-9]*\.?[0-9]+)-([-+]?[0-9]*\.?[0-9]+)[\s-](\d*)')
    
#   trapped_tag = re.compile('Trapped - Zone')
    escaped_tag = re.compile('Escaped - Zone')
    basicInfo = [] # diameter and angle
#    trapped_temp = {}    
    escaped_temp = {}
    
#   extract 'injection-75-141' style expression    
    for i in range(np.size(chunk)):
        # find 'Escaped - Zone' line
        if (escaped_tag.search(chunk[i]) != None):
            # find injection information 'injection-75-142.9  198'
            if(r_tag.search(chunk[i]) != None):
                # This is the injection info section
                # extract the numbers of the 'injection...' string
                r_list  = {np.float(x) for x in r_tag.findall(chunk[i])[0]}
                r       = r_list[1]
                basicInfo.append(r_list[0])
                basicInfo.append(r_list[2])
            else:
                # This is the injection results section
                stats = map(float, sn_tag.findall(chunk[i]))
                stats[0] = int(stats[0])
                escaped_temp[stats[0]] = stats[1:]
                host[r] = escaped_temp
    return basicInfo 
    
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
