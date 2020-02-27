#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Script to bulk convert CORD Stat Act Comparison Reports into much more
usefully fomatted Excel spreadsheets.

The Stat Act Comparison Report zip files should be placed in the INPUT folder.
Run the script, and the resulting spreadsheet can be found in the OUTPUT
folder.

Script can currently only create one OUTPUT file for ONE comparison (eg INT19-
Bluebook). If other comparison zip's are found in the input file (eg DEV19-
Bluebook), the file will be skipped.
"""

import os, sys, glob
import numpy as np
import pandas as pd
from datetime import datetime
import CORDtools as ct

def compileCO(summaryDf, idxName):
    """Compiles summary dataframe's into coDf, where each summaryDf is a row.
    """
    global coDf
    for i, column in enumerate(summaryDf.loc[0,:]):
        coDf.loc[idxName, column] = summaryDf.loc[1, i]
    coDf = coDf.replace('N', np.nan)

def readSummmary(dirname):
    """Reads the Comparison Report file within specified directory (dirname).
    Returns the summary file dataframe, name of the system, the source and 
    target modes.
    
    Skips first 5 rows of the file and transposes the datframe to create
    summaryDf. headerDf is also created which is only used the read the file
    header which contains the source and target of the comparison for later
    checks.
    """
    global inpFol
    os.chdir(dirname)
    files = []
    for (dirpath, dirnames, filenames) in os.walk(os.getcwd()):
        files.extend(filenames)
        break
    for file in files:
        if '_Comparison_Report_Summary_' in file:
            splitDir = dirname.split('_Stat_Act_Comparision_Report_')
            summaryDf = pd.read_csv(file, skiprows=5, header=None,
                                       encoding='unicode_escape')
            headerDf = pd.read_csv(file, header=None, sep='|',
                                       encoding='unicode_escape')
            infoLine = headerDf.loc[2, 0]
            splitInfo = infoLine.split('/')
            src = splitInfo[1].split(' ')[0]
            trg = splitInfo[2]
            summaryDf = summaryDf.T
        #No_objects_found.txt fills report summary files when a Stat Act Comp
        #is run on an empty Statistical Activity.
        if 'No_objects_found' in file:
            summaryDf = pd.DataFrame()
            src = 'invalid'
            trg = 'invalid'
            splitDir = ['invalid']
    os.chdir(inpFol)
    return summaryDf, splitDir[0], src, trg
    
def formatComparison(writer):
    """Formats the Comparison Overview worksheet.
    
    Sets column widths and assigns relevant formatting to cells.
    """
    global coDf
    wb = writer.book
    ws = writer.sheets['Comparison Overview']
    r = len(coDf.iloc[:, 0]) + 1
    diff_format = wb.add_format({'bg_color': '#f67c7c',
                                 'border_color': '#d60000',
                                 'border': 5,
                                 'align': 'center',
                                 'valign': 'vcenter',
                                 'bold': True,
                                 'font_color': 'white'})
    empty_format = wb.add_format({'bg_color': 'white',
                                 'border': 1,
                                 'align': 'center',
                                 'valign': 'vcenter'})
    diffDf = coDf.replace(np.nan, '')
    for c in range(len(coDf.iloc[0, :])):
        for r in range(len(coDf.iloc[:, 0])):
            if diffDf.iloc[r, c] != '':
                ws.write(r+1, c+1, 'Different', diff_format)
            else:
                ws.write(r+1, c+1, '', empty_format)
    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 14)
    ws.set_column(2, 2, 10)
    ws.set_column(3, 3, 10)
    ws.set_column(4, 4, 11)
    ws.set_column(5, 5, 13)
    ws.set_column(6, 6, 12)
    ws.set_column(7, 7, 12)
    ws.set_column(8, 8, 10)
    ws.set_column(9, 9, 16)
    ws.set_column(10, 10, 15)
    ws.set_column(11, 11, 12)

def adjustCO(sys, field):
    """Adjusts the Comparison Overview sheet in the case a badly encoded csv
    is found.
    """
    global writer
    global coDf
    returnPath = os.getcwd()
    os.chdir(outFol)
    wb = writer.book
    ws = writer.sheets['Comparison Overview']
    col = None
    for c, title in enumerate(coDf.columns):
        if title == field:
            col = c
    if col == None:
        ct.error('Field not found!')
    row = None
    for r, idx in enumerate(coDf.index):
        if idx == sys:
            row = r
    if row == None:
        ct.error('System not found!')
    ws.write(row+1, col+1, 'Review', wb.add_format({'border': 5,
                                                  'bg_color': 'red',
                                                  'border_color': '#cc3300',
                                                  'align': 'center',
                                                  'valign': 'vcenter',
                                                  'bold': True,
                                                  'font_color': 'white'}))
    os.chdir(returnPath)

def createSheetDf(sheet):
    """Returns the dataframe holding the comparison field data for all systems
    where the comparison report showed differences.
    """
    global curDir
    files = []
    sheetFileName = None
    #Convert column header (sheet) to filename (sheetToFile) based on
    #sheetDict
    for sheetToFile in sheetDict:
        if sheet == sheetToFile[0]:
            sheetFileName = sheetToFile[1]
    if sheetFileName == None:
        ct.error('Unrecognised sheet, skipping...')
        return
    #Create a list containing all files
    for (dirpath, dirnames, filenames) in os.walk(os.getcwd()):
        files.extend(filenames)
        break
    #loop through files until we find the one we're looking for
    for file in files:
        if sheetFileName in file:
            dataStart = None
            basAtr = None
            #Try to read the csv, if the encoding is unrecognised the send
            #it to adjustCO to show it's been skipped on output sheet.
            try:
                prelimDf = pd.read_csv(file, header=None, sep='\n',
                               encoding = 'unicode_escape')
                #prelimDf = pd.read_csv(file, header=None, sep='\n', 
                 #                      ct.error_bad_lines=False)
            except:
                ct.error('BAD ENCODING IN '+curDir+' \nSKIPPING DATA,'+
                      'THIS WILL NEED TO BE REVIEWED MANUALLY...\033[93m')
                adjustCO(curDir.split('_')[0], sheet)
                continue
            for i, line in enumerate(prelimDf.iloc[:,0]):
                #Find where the data starts
                if 'Name,Match' in line:
                    dataStart = i+2
                    #Added placeholder to account for final commas
                    headerNames = line.split(',')+['Placeholder']
                #Grab the basic attribute data as it's variable between files
                if 'Basic Attributes: ' in line:
                    basAtr = line.split('Basic Attributes: ')[-1]
                    basAtr = basAtr[:-2]
            #Get the sheet dataframe containing it's predefined columns
            sheetDf = pd.read_csv(file, skiprows=dataStart, quotechar='"',
                               names=headerNames, encoding = 'unicode_escape')
            sheetDf.drop('Placeholder', inplace=True, axis=1)
            #Create a dataframe to hold data from all files
            finalDf = pd.DataFrame(columns=['System',
                                            'Name',
                                            'Location',
                                            'Differences'])
            #Loop through all rows of sheetDf and append the rows that showed
            #differences to the finalDf, listing their differences in the
            #Differences column
            for r, item in enumerate(sheetDf.loc[:, 'Identical']):
                rowDf = pd.DataFrame()
                if item != 'Y':
                    diffStr = ''
                    rowDf.loc[0, 'System'] = curDir.split('_')[0]
                    if  not pd.notnull(item) and not\
                        pd.notnull(sheetDf.loc[r,'Match']):
                        rowDf.loc[0, 'Name'] = sheetDf.loc[r, 'Items']
                        if sheetDf.loc[r, 'Groups']=='Descriptions different':
                            diffStr = sheetDf.loc[r, 'Groups']
                            rowDf.loc[0, 'Location'] = 'In Both'
                        else:
                            rowDf.loc[0, 'Location'] = sheetDf.loc[r, 'Groups']
                    else:
                        rowDf.loc[0, 'Name'] = sheetDf.loc[r, 'Name']
                        rowDf.loc[0, 'Location'] = sheetDf.loc[r, 'Match']
                    diffs = []
                    if item == 'N':
                        for i, x in enumerate(sheetDf.iloc[r, 3:]):
                            if pd.notnull(x):
                                colHead = sheetDf.columns[i+3]
                                if colHead == 'Basic Attributes':
                                    diffs.append('Basic Attributes (' + 
                                                 basAtr + ')')
                                else:
                                    diffs.append(colHead)
                    if diffs != []:
                        diffStr = (', ').join(diffs)
                    rowDf.loc[0, 'Differences'] = diffStr
                    #Final check to make sure data isn't just misplaced headers
                    if rowDf.loc[0, 'Name'] == 'Name' and \
                    rowDf.loc[0, 'Location'] == 'Match':
                        continue
                    finalDf = finalDf.append(rowDf, ignore_index=True)
                    
            return finalDf

def createSheet(sheet):
    """Creates and formats the sheet for a specific comparison field (sheet).
    """
    global writer
    global inpFol
    global outFol
    global dirs
    global coDf
    global sheetDict
    global curDir
    os.chdir(inpFol)
    changed = []
    dfList = []
    print('Creating sheet for ' + sheet + '...')
    #Loop through all systems in comparison overview dataframe to find those
    #with differences
    for i, item in enumerate(coDf.loc[:, sheet]):
        #If system has changed for this comparison field append it to the list
        if item == 'Y':
            changed.append(coDf.index[i])
    #Create a list of dataframes for each system that has changes in the
    #comparison field (sheet)
    for changedSys in changed:
        for aDir in dirs:
            if changedSys in aDir:
                os.chdir(inpFol)
                os.chdir(aDir)
                curDir = aDir
                dfList.append(createSheetDf(sheet))
    bigDf = pd.DataFrame()
    #Compile the list of dataframes into one big dataframe
    for df in dfList:
        df = df.replace('',np.nan)
        df.dropna(how='all', subset=['Name','Location','Differences'],
                  inplace=True)
        df = df.replace(np.nan, '')
        bigDf = bigDf.append(df)
    #Create a worksheet using xlsxwriter
    wb = writer.book
    ws = wb.add_worksheet(sheet)
    #Create variable to hold the system name, as when a system changes we
    #want a larger border to visually seperate systems
    prevSys = None
    #Loop through each row in the bigDf and each column item to the sheet in
    #the proper formatting
    for r, system in enumerate(bigDf.loc[:, 'System']):
        std_format = wb.add_format({'bg_color': 'white',
                                    'border': 1,
                                    'align': 'left',
                                    'valign': 'vcenter'})
        sys_format = wb.add_format({'bg_color': 'white',
                                    'border': 1,
                                    'right': 5,
                                    'left': 5,
                                    'align': 'center',
                                    'bold': True,
                                    'valign': 'vcenter'})
        if r == 0:
            ws.write(0,0, 'System', wb.add_format({'border': 5,
                                                   'bold': True,
                                                   'align': 'center',
                                                   'valign': 'vcenter'}))
            ws.write(0,1, 'Name', wb.add_format({'top': 5,
                                                 'bottom': 5,
                                                 'left': 5,
                                                 'right': 1,
                                                 'bold': True,
                                                 'align': 'center',
                                                 'valign': 'vcenter'}))
            ws.write(0,2, 'Location', wb.add_format({'top': 5,
                                                 'bottom': 5,
                                                 'left': 1,
                                                 'right': 1,
                                                 'bold': True,
                                                 'align': 'center',
                                                 'valign': 'vcenter'}))
            ws.write(0,3, 'Differences', wb.add_format({'top': 5,
                                                 'bottom': 5,
                                                 'left': 1,
                                                 'right': 5,
                                                 'bold': True,
                                                 'align': 'center',
                                                 'valign': 'vcenter'}))
        #Sets up the thick borders for system separation
        if system != prevSys:
            std_format.set_top(5)
            sys_format.set_top(5)
            prevSys = system
        
        #Sets a thick border for the final line of the dataset
        if r == len(bigDf.iloc[:, 0])-1:
            std_format.set_bottom(5)
            sys_format.set_bottom(5)
        
        #Writes the data to the sheet for each item in each column
        for c, info in enumerate(bigDf.iloc[r, :]):
            if c == 0:
                form = sys_format
            else:
                form = std_format
            ws.write(r+1, c, info, form)
        #Get the cell range of the dataset as a string
        cellStr = 'E2:E' + str(len(bigDf.iloc[:,0])+1)
        #Make the far right border of the dataset thick
        ws.conditional_format(cellStr, {'type': 'blanks',
                                   'format': wb.add_format({'left': 5})
                                   })
                                        
    #Set the column widths
    ws.set_column(0, 0, 27)
    ws.set_column(1, 1, 40)
    ws.set_column(2, 2, 11)
    ws.set_column(3, 3, 126)
        
def runTask():
    """Runs the overall task, converting zip files in the INPUT folder into
    an xlsx sheet in the OUTPUT folder.
    """
    global inpFol
    global outFol
    global coDf
    global writer
    global dirs
    origSrc = None
    origTrg = None
    (inpFol, outFol) = ct.setupFilepaths()
    os.chdir(inpFol)
    dirs = []
    #Unzip all zip files
    if list(glob.glob('*.zip')) == []:
        ct.error('NO INPUT FILES FOUND. Place CORD Stat Act Comparison zip '+
              'files in the INPUT file and re-run.')
        sys.exit()
    for file in glob.glob('*.zip'):
        if 'Stat_Act_Comparision_Report' not in file:
            ct.error(file + ' is not a Comparison Report, this file will be '+
                  'skipped.', warning=True)
            continue
        skip = False
        for stat in dirs:
            cur = file.split('_Stat_Act_Comparision_Report_')[0]
            if stat.split('_Stat_Act_Comparision_Report_')[0] == cur:
                ct.error('There are multiple Reports for the Stat Act: ' + cur +
                      '. Only the file "' + stat + '" will be analysed.',
                      warning=True)
                skip=True
        if skip: continue
        ct.unzipFiles(file)
        dirs.append(os.path.splitext(file)[0])
    #Read all directories to create the comparison dataset
    for aDir in dirs:
        print('Reading summary for', aDir, '...')
        sList, idxName, src, trg = readSummmary(aDir)
        #If the zip file contains no files it will return 'invalid'
        #If invalid is found, throw and error and skip the folder
        if 'invalid' in {src, trg}:
            ct.error('This report contains no information! Skipping...')
            continue
        if(origSrc == None and origTrg == None):
            origSrc = src
            origTrg = trg
        if src == origSrc and trg == origTrg:
            print('Adding to comparison...')
            compileCO(sList, idxName)
        else:
            ct.error('Source and target were different!' + 
                  ' Skipping comparison...')
            
    print('Creating Comparison Overview...')
    date = datetime.today().strftime('%d-%m-%Y')
    compTitle = 'Comparison Overview (' + src + '-' + trg + ') ' + date +\
    '.xlsx'
    sameFiles = []
    for file in glob.glob('*.xlsx'):
        if compTitle[:-5] in file:
            sameFiles.append(file)
    if os.path.exists(compTitle):
        compTitle = compTitle[:-5] + '(' + str(len(sameFiles)) + ').xlsx'
    writer = pd.ExcelWriter(compTitle, engine='xlsxwriter')
    coDf.columns = ['CLASSIFICATION',
                    'TASK',
                    'MAPPINGS',
                    'PARAMETER',
                    'CALCULATIONS',
                    'CON CHECKS',
                    'IMPORT DEFS',
                    'DATASETS',
                    'SEAS EXPORT DEFS',
                    'VISUALISATIONS',
                    'EXPORT DEFS']
    coDf.to_excel(writer, sheet_name='Comparison Overview')
    print('Applying formatting to Comparison Overview...')
    formatComparison(writer)
    
    #Create a comparison sheet for every comparison field that has a changed
    #system
    coDfAdj = coDf.dropna(how='all', axis=1)
    for col in coDfAdj.columns:
        createSheet(col)
            
    #Save the file
    os.chdir(outFol)
    writer.save()
#Predefine the comparison overview dataframe to be accessible to all functions
coDf = pd.DataFrame(columns = ['CLASSIFICATION',
                               'TASK',
                               'CLASSIFICATION MAPPING',
                               'PARAMETER',
                               'CALCULATION DEFINITION',
                               'CONSISTENCY CHECK DEFINITION',
                               'TIMESERIES IMPORT DEFINITION',
                               'DATASET DEFINITION',
                               'SEAS EXPORT DEFINITION',
                               'VISUALISATION DEFINITION',
                               'TIMESERIES EXPORT DEFINITION'])
#Predefine the sheet dictionary which helps read the csv files
sheetDict = [['CLASSIFICATION', 'Classifications'],
             ['TASK', 'Tasks'],
             ['PARAMETER', 'Parameters'],
             ['MAPPINGS', 'Classification_Mappings'],
             ['CALCULATIONS', 'Calculations'],
             ['CON CHECKS', 'Consistency_Checks'],
             ['IMPORT DEFS', 'Import_Definitions'],
             ['DATASETS', 'Dataset_Definitions'],
             ['SEAS EXPORT DEFS', 'Seas_Analysis_Export_Definitions'],
             ['VISUALISATIONS', 'Visualisations'],
             ['EXPORT DEFS', 'Export_Definitions']]
             
#Run the script
runTask()
ct.done()