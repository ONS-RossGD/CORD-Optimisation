# -*- coding: utf-8 -*-
"""Script to search CORD Statistical Activities for redundancies.

A CORD Config Report should be downloaded for each Statistical Activity that
you'd like to test for redundancies and placed in the INPUT file. A
Redundancy Report will be created for each Stat Act where redundancies are 
found and placed in a folder of the date and time the script was run inside
of the OUTPUT folder.

If the Stat Act has external dependencies, these will be accounted for as long
as Config Reports are also provided for said external dependencies. If the
Config Reports are not provided a warning will appear to notify you.

The script currently checks for the following:
    - Flagged Tasks: Tasks that have not been run in over 1 year.
    - Redundant Calcs: Calculations that are not in any Task so will never be
                       run.
    - Redundant Parameters: Parameters that are never used.
    - Flagged Class Groups: Classification Item Groups that are not used in
                            any Calc or Consistency Check within Stat Acts in
                            the INPUT file.
    - Flagged Visualisations: Visualisations that haven't been used in over 1
                              year.
    - Redundant Con Checks: Consistency Checks that are not in any Task so will
                            never get run.
    - Redundant Imports: Import Definitions that  are not in any Task so will
                         never get run.
    - Redundant Exports: Export Definitions that  are not in any Task so will
                         never get run.
Author: Ross Gregory-Davies : gregor1
"""
import pandas as pd
import glob, os, zipfile, datetime
import CORDtools as ct

def unzipFiles():
    """Unzips all zip files in the input folder.
    """
    global configReports
    os.chdir(inpFile)
    for file in glob.glob('*.zip'):
        if '_Config_Report_' not in file:
            ct.error('"'+file+'" is not a Config Report and will be ignored.',
                  warning=True)
            continue
        fileTitle = os.path.splitext(file)[0]
        splitTitle = fileTitle.split('_Config_Report_')
        r = configReports['Config Report'].count()
        configReports.loc[r, 'Config Report'] = fileTitle
        configReports.loc[r, 'Statistical Activity'] = splitTitle[0]
        configReports.loc[r, 'Mode'] = splitTitle[1].split('_')[0]
        configReports.loc[r, 'Date'] = ' '.join(splitTitle[1].split('_')[1:])
        print('Unzipping', fileTitle, '...')
        with zipfile.ZipFile(file, 'r') as zipObj:
             zipObj.extractall(fileTitle)

def createOutFolder():
    global outFolder
    """Changes directories to the output folder setting one up if it doesn't 
    already exist.
    """
    returnPath = os.getcwd()
    os.chdir(outFile)
    fol = datetime.datetime.now().strftime("%d-%m-%Y (%H;%M;%S)")
    try:
        os.chdir(fol)
    except:
        os.mkdir(fol)
        os.chdir(fol)
    outFolder = os.getcwd()
    os.chdir(returnPath)

def readTasks():
    global tasksDf
    tasksDf = pd.DataFrame()
    for file in glob.glob('*.csv'):
        if 'Tasks_rpt_' in file:
            tasksDf = pd.read_csv(file, skiprows=3, encoding='unicode_escape')
            for r, usedDate in enumerate(tasksDf['Date Last Used']):
                tasksDf.loc[r, 'Date Last Used'] = datetime.datetime.strptime(\
                           tasksDf.loc[r,'Date Last Used'],'%d/%m/%Y %H:%M:%S')
            tasksDf.drop('Is Parallel', axis=1, inplace=True)
            tasksDf = tasksDf.set_index('Name')

def replaceTaskDate(parent, child):
    global tasksDf
    if parent not in tasksDf.index.tolist():
        return
    if child not in tasksDf.index.tolist():
        return    
    parentUsed = tasksDf.loc[parent,'Date Last Used']
    childUsed = tasksDf.loc[child,'Date Last Used']
    if parentUsed > childUsed:
        tasksDf.loc[child, 'Date Last Used'] = parentUsed

def readTaskLns():
    global usedCalcs
    global usedImports
    global usedExports
    global usedConChecks
    tempDf = pd.DataFrame()
    hasTasks = False
    usedCalcs = []
    for file in glob.glob('*.csv'):
        if 'TskLn-' in file:
            hasTasks = True
            prelimDf = pd.read_csv(file, sep='|', header=None,
                                   encoding='unicode_escape')
            parent = prelimDf.loc[1, 0].split('Task: ')[1]\
                     .replace(' (Not Parallel)', '').replace(' (Parallel)', '')
            taskLnDf = pd.read_csv(file, skiprows=3, encoding='unicode_escape')
            taskIdxs = taskLnDf[taskLnDf['Type']=='TASK'].index.tolist()
            for child in taskLnDf.loc[taskIdxs, 'Name']:
                replaceTaskDate(parent, child)
            tempDf = pd.concat([tempDf,taskLnDf])
    if hasTasks:
        tempDf.reset_index(inplace=True, drop=True)
        calcIdxs = tempDf[tempDf['Type']=='CALCULATION DEFINITION']\
        .index.tolist()
        importsIdxs = tempDf[tempDf['Type']=='TIMESERIES IMPORT DEFINITION']\
                        .index.tolist()
        exportIdxs = tempDf[tempDf['Type']=='TIMESERIES EXPORT DEFINITION']\
                        .index.tolist()
        conCheckIdxs = tempDf[tempDf['Type']=='CONSISTENCY CHECK DEFINITION']\
                        .index.tolist()
        usedCalcs = tempDf.loc[calcIdxs, 'Name'].drop_duplicates().tolist()
        usedImports = tempDf.loc[importsIdxs, 'Name'].drop_duplicates()\
                        .tolist()
        usedExports= tempDf.loc[exportIdxs, 'Name'].drop_duplicates()\
                        .tolist()
        usedConChecks = tempDf.loc[conCheckIdxs, 'Name'].drop_duplicates()\
                        .tolist()

def readCalcs():
    global calculationsDf
    calculationsDf = pd.DataFrame()
    for file in glob.glob('*.csv'):
        if 'Calculations_rpt_' not in file: continue
        calculationsDf = pd.read_csv(file, skiprows=3,
                                     encoding='unicode_escape')
        
def readConChecks():
    global conCheckDf
    conCheckDf = pd.DataFrame()
    for file in glob.glob('*.csv'):
        if file[:22] != 'Consistency_Checks_rpt': continue
        conCheckDf = pd.read_csv(file, skiprows=3, encoding='unicode_escape')
        for r in range(len(conCheckDf.index.tolist())):
            conCheckDf.loc[r, 'Date Last Used'] = datetime.datetime.strptime(\
                           conCheckDf.loc[r,'Date Last Used'],
                           '%d/%m/%Y %H:%M:%S')
              
def readImports():
    global importsDf
    importsDf = pd.DataFrame()
    for file in glob.glob('*.csv'):
        if file[:23] != 'Import_Definitions_rpt_': continue
        importsDf = pd.read_csv(file, skiprows=3, encoding='unicode_escape')
        for r in range(len(importsDf.index.tolist())):
            importsDf.loc[r, 'Date Last Used'] = datetime.datetime.strptime(\
                           importsDf.loc[r,'Date Last Used'],
                           '%d/%m/%Y %H:%M:%S')
            importsDf = importsDf[['Name', 'Description','Date Created',
                           'Date Last Used']]
              
def readExports():
    global exportsDf
    exportsDf = pd.DataFrame()
    for file in glob.glob('*.csv'):
        if file[:23] != 'Export_Definitions_rpt_': continue
        exportsDf = pd.read_csv(file, skiprows=3, encoding='unicode_escape')
        for r in range(len(exportsDf.index.tolist())):
            exportsDf.loc[r, 'Date Last Used'] = datetime.datetime.strptime(\
                           exportsDf.loc[r,'Date Last Used'],
                           '%d/%m/%Y %H:%M:%S')
            exportsDf = exportsDf[['Name', 'Description','Date Created',
                           'Date Last Used']]
        
def readVisualisations():
    global visDf
    visDf = pd.DataFrame()
    for file in glob.glob('*.csv'):
        if 'Visualisations_rpt_' not in file: continue
        visDf = pd.read_csv(file, skiprows=3,
                                     encoding='unicode_escape')
        for r in range(len(visDf.index.tolist())):
            visDf.loc[r, 'Date Last Used'] = datetime.datetime.strptime(\
                           visDf.loc[r,'Date Last Used'],'%d/%m/%Y %H:%M:%S')
            visDf = visDf[['Name', 'Description','Date Created',
                           'Date Last Used']]

def readParams():
    global paramDf
    paramDf = pd.DataFrame()
    for file in glob.glob('*.csv'):
        if 'Parameters_rpt_' not in file: continue
        paramDf = pd.read_csv(file, skiprows=3, encoding='unicode_escape')
            
def readClassGrps():
    global classGrpsDf
    classGrpDf = pd.DataFrame()
    classGrpsDf = pd.DataFrame(columns=['Name','Description','Classification'])
    for file in glob.glob('*.csv'):
        if 'Grps' in file:
            prelimDf = pd.read_csv(file, sep='|', header=None,
                                   encoding='unicode_escape')
            clas = prelimDf.loc[1,0].split('Classification: ')[1]
            classGrpDf = pd.read_csv(file, skiprows=3, 
                                     encoding='unicode_escape')
            classGrpDf.loc[:, 'Classification'] = clas
            classGrpsDf = pd.concat([classGrpsDf,
                                     classGrpDf.loc[:, ['Name','Description',
                                                        'Classification']]])

def readExtClass():
    global extClassDf
    extClassDf = pd.DataFrame()
    for file in glob.glob('*.csv'):
        if 'External_Classifications_Out_' in file:
            extClassDf = pd.read_csv(file, skiprows=4,
                                     encoding='unicode_escape')

def readExtDatasets():
    global extDatasetDf
    extDatasetDf = pd.DataFrame()
    for file in glob.glob('*.csv'):
        if 'External_Datasets_Out_' in file:
            extDatasetDf = pd.read_csv(file, skiprows=3,
                                     encoding='unicode_escape')
    
def findRedundantCalcs():
    global redundantCalcs
    global calculationsDf
    global calcsMinusRedundantDf
    redundantCalcs = list(set(usedCalcs)^set(calculationsDf['Name'].tolist()))
    calcsMinusRedundantDf = calculationsDf.copy()
    calculationsDf = calculationsDf.set_index('Name')
    for calc in redundantCalcs:
        calcsMinusRedundantDf=calcsMinusRedundantDf[~calcsMinusRedundantDf\
                                                    ['Name'].str.match(calc)]

def findRedundantConChecks():
    global redundantConChecks
    global conCheckDf
    global conCheckMinusRedundantDf
    redundantConChecks = list(set(usedConChecks)^set(conCheckDf['Name']\
                              .tolist()))
    conCheckMinusRedundantDf = conCheckDf.copy()
    conCheckDf = conCheckDf.set_index('Name')
    for conCheck in redundantConChecks:
        conCheckMinusRedundantDf=conCheckMinusRedundantDf\
                                 [~conCheckMinusRedundantDf\
                                 ['Name'].str.match(conCheck)]

def findRedundantImports():
    global redundantImports
    global importsDf
    redundantImports = list(set(usedImports)^set(importsDf['Name']\
                              .tolist()))
    importsDf = importsDf.set_index('Name')

def findRedundantExports():
    global redundantExports
    global exportsDf
    redundantExports = list(set(usedExports)^set(exportsDf['Name']\
                              .tolist()))
    exportsDf = exportsDf.set_index('Name')

def findRedundantParams():
    global redundantParams
    global paramDf
    for p, param in enumerate(paramDf['Name']):
        startDates = calcsMinusRedundantDf['Start Date'].tolist() +\
                        conCheckMinusRedundantDf['Start Date'].tolist()
                        
        while ' ' in startDates:
            startDates.remove(' ')
        startDates = list(dict.fromkeys(startDates))
        startDates = [sd[1:] for sd in startDates]
        endDates = calcsMinusRedundantDf['End Date'].tolist() +\
                        conCheckMinusRedundantDf['End Date'].tolist()
        while ' ' in endDates:
            endDates.remove(' ')
        endDates = list(dict.fromkeys(endDates))
        endDates = [ed[1:] for ed in endDates]
        
        if param in startDates or param in endDates:
            paramDf.drop(p, inplace=True)
    if paramDf.empty: return
    formulaCalcs = calcsMinusRedundantDf[calcsMinusRedundantDf['Type']\
                                         .str.match('FORMULA')]
    for typeSpecDet in formulaCalcs['Type Specific Details']:
        formula = typeSpecDet.split('{formula = ')[1]
        formula = formula.split('}')[0]
        for char in {',','+','-','/',')','(','{','}','*'}:
            if char in formula:
                formula = formula.replace(char, ' ')
        formula = formula.split('$')
        if len(formula) > 1:
            for i in range(1, len(formula)):
                param = formula[i].split(' ')[0].strip()
                if param in paramDf['Name'].tolist():
                    paramDf = paramDf[~paramDf['Name'].str.match(param)]
    redundantParams = paramDf['Name'].tolist()

def getMassiveStr(calcDf=pd.DataFrame(), imDf=pd.DataFrame(),
                  exDf=pd.DataFrame(), ccDf=pd.DataFrame()):
    massiveStr = ''
    if not calcDf.empty:
        massiveStr += ' '.join(calcDf['Type Specific Details']\
                              .tolist())
        massiveStr += ' '.join(calcDf['Selection Criteria']\
                               .tolist())
    if not imDf.empty:
        massiveStr += ' '.join(imDf['Type Specific Details']\
                              .tolist())
        massiveStr += ' '.join(imDf['Selection Criteria']\
                               .tolist())
    if not exDf.empty:
        massiveStr += ' '.join(exDf['Type Specific Details']\
                              .tolist())
        massiveStr += ' '.join(exDf['Selection Criteria']\
                               .tolist())
    if not ccDf.empty:
        massiveStr += ' '.join(ccDf['Type Specific Details']\
                              .tolist())
        massiveStr += ' '.join(ccDf['Selection Criteria']\
                               .tolist())
    return massiveStr

def searchClassGrpsDependencies():
    global usedGrpsExt
    global grpSrchDf
    hiddenStatActs = []
    for i, statAct in enumerate(extClassDf['Statistical Activity'].tolist()):
        if statAct not in configReports['Statistical Activity'].tolist():
            hiddenStatActs.append(statAct)
            extClassDf.drop(i, inplace=True)
            continue
        if statAct not in grpSrchDf.columns:
            os.chdir(inpFile)
            fol = configReports.loc[configReports['Statistical Activity']\
                          .isin([statAct]), 'Config Report'].tolist()[0]
            os.chdir(fol)
            for file in glob.glob('*.csv'):
                imDf = pd.DataFrame()
                exDf = pd.DataFrame()
                ccDf = pd.DataFrame()
                if file[:17] == 'Calculations_rpt_':
                    calcDf = pd.read_csv(file, skiprows=3,
                                         encoding='unicode_escape')
                if file[:23] == 'Export_Definitions_rpt_':
                    exDf = pd.read_csv(file, skiprows=3,
                                         encoding='unicode_escape')
                if file[:23] == 'Import_Definitions_rpt_':
                    imDf = pd.read_csv(file, skiprows=3,
                                         encoding='unicode_escape')
                if file[:23] == 'Consistency_Checks_rpt_':
                    ccDf = pd.read_csv(file, skiprows=3,
                                         encoding='unicode_escape')
                massiveStr = getMassiveStr(calcDf=calcDf, imDf=imDf, exDf=exDf,
                                           ccDf=ccDf)
            grpSrchDf.loc[0, statAct] = massiveStr
        else:
            massiveStr = grpSrchDf.loc[0, statAct]
        for r, grp in enumerate(classGrpsDf['Name']):
            if 'group '+grp in massiveStr:
                usedGrpsExt.append(grp)
    hiddenStatActs = list(dict.fromkeys(hiddenStatActs))
    if len(hiddenStatActs) > 0:
        ct.error(curStatAct+' has external classification dependancies '+
             "that aren't contained in the INPUT file.", warning=True)

def findRedundantClassGrps():
    global usedGrps
    global grpSrchDf
    massiveStr = getMassiveStr(calcDf=calcsMinusRedundantDf, exDf=extDatasetDf,
                                           ccDf=conCheckMinusRedundantDf)
    grpSrchDf.loc[0, curStatAct] = massiveStr
    for r, grp in enumerate(classGrpsDf['Name']):
        if 'group '+grp in massiveStr:
            usedGrps.append(grp)

def flagTasks():
    global flaggedTasks
    for r, taskUsedDate in enumerate(tasksDf['Date Last Used']):
        if taskUsedDate < timeTol:
            flaggedTasks.append(tasksDf.index[r])

def flagVisualisations():
    global flaggedVis
    for r, visUsedDate in enumerate(visDf['Date Last Used']):
        if visUsedDate < timeTol:
            flaggedVis.append(visDf.index[r])
        

def compileUsedClassGrps():
    global flaggedGrps
    usedGrpsTot = usedGrps+usedGrpsExt
    flaggedGrps = list(set(usedGrpsTot)^set(classGrpsDf['Name'].tolist()))

def formatSheet(sheet, df):
    wb = writer.book
    ws = writer.sheets[sheet]
    center_format = wb.add_format({'align': 'center',
                                   'valign': 'center'})
    if sheet == 'Flagged Tasks':
        style = 'Table Style Medium 7'
        ws.add_table(0,0,len(df.index),3, ct.createFormat(df, style,idx=True))
        ws.set_column('A:A', 35, center_format)
        ws.set_column('B:B', 129, center_format)
        ws.set_column('C:C', 20, center_format)
        ws.set_column('D:D', 20, center_format)
    if sheet == 'Redundant Calcs':
        style = 'Table Style Medium 3'
        ws.add_table(0,0,len(df.index),3, ct.createFormat(df, style,idx=True))
        ws.set_column('A:A', 35, center_format)
        ws.set_column('B:B', 129, center_format)
        ws.set_column('C:C', 20, center_format)
        ws.set_column('D:D', 20, center_format)
    if sheet == 'Redundant Parameters':
        style = 'Table Style Medium 2'
        ws.add_table(0,0,len(df.index),3, ct.createFormat(df, style))
        ws.set_column('A:A', 35, center_format)
        ws.set_column('B:B', 129, center_format)
        ws.set_column('C:C', 20, center_format)
        ws.set_column('D:D', 20, center_format)
    if sheet == 'Flagged Class Groups':
        style = 'Table Style Medium 5'
        ws.add_table(0,0,len(df.index),2, ct.createFormat(df, style))
        ws.set_column('A:A', 40, center_format)
        ws.set_column('B:B', 130, center_format)
        ws.set_column('C:C', 35, center_format)
    if sheet == 'Flagged Visualisations':
        style = 'Table Style Medium 6'
        ws.add_table(0,0,len(df.index),3, ct.createFormat(df, style))
        ws.set_column('A:A', 35, center_format)
        ws.set_column('B:B', 129, center_format)
        ws.set_column('C:C', 20, center_format)
        ws.set_column('D:D', 20, center_format)
    if sheet == 'Redundant Con Checks':
        style = 'Table Style Medium 1'
        ws.add_table(0,0,len(df.index),3, ct.createFormat(df, style))
        ws.set_column('A:A', 35, center_format)
        ws.set_column('B:B', 129, center_format)
        ws.set_column('C:C', 20, center_format)
        ws.set_column('D:D', 20, center_format)
    if sheet == 'Redundant Imports':
        style = 'Table Style Medium 4'
        ws.add_table(0,0,len(df.index),3, ct.createFormat(df, style))
        ws.set_column('A:A', 35, center_format)
        ws.set_column('B:B', 129, center_format)
        ws.set_column('C:C', 20, center_format)
        ws.set_column('D:D', 20, center_format)
    if sheet == 'Redundant Exports':
        style = 'Table Style Medium 4'
        ws.add_table(0,0,len(df.index),3, ct.createFormat(df, style))
        ws.set_column('A:A', 35, center_format)
        ws.set_column('B:B', 129, center_format)
        ws.set_column('C:C', 20, center_format)
        ws.set_column('D:D', 20, center_format)
        
def write():
    global writer
    filled = False
    os.chdir(outFolder)
    if len(flaggedTasks)+len(redundantCalcs)+len(redundantParams)\
            +len(flaggedGrps)+len(flaggedVis)+len(redundantConChecks)\
            +len(redundantImports)+len(redundantExports)>0:
        writer = pd.ExcelWriter(curStatAct + ' Redundancy Report.xlsx', 
                            engine='xlsxwriter')
    if len(flaggedTasks) > 0:
        filled = True
        flaggedTaskSheetDf = tasksDf.loc[flaggedTasks].copy()
        for r, dt in enumerate(flaggedTaskSheetDf['Date Last Used']):
            flaggedTaskSheetDf.loc[flaggedTaskSheetDf.index[r],
                                   'Date Last Used']=\
            datetime.datetime.strftime(dt, '%d/%m/%Y %H:%M:%S')
        flaggedTaskSheetDf.to_excel(writer, sheet_name='Flagged Tasks')
        formatSheet('Flagged Tasks', flaggedTaskSheetDf)
    if len(redundantCalcs) > 0:
        filled = True
        redundantCalcSheetDf = calculationsDf.loc[redundantCalcs,
                                                 ['Description','Date Created',
                                                  'Date Last Used']].copy()
        redundantCalcSheetDf.to_excel(writer, sheet_name='Redundant Calcs')
        formatSheet('Redundant Calcs', redundantCalcSheetDf)
    if len(redundantParams) > 0:
        filled = True
        redundantParamSheetDf = paramDf.copy()
        redundantParamSheetDf.to_excel(writer,
                                       sheet_name='Redundant Parameters',
                                       index=False)
        formatSheet('Redundant Parameters', redundantParamSheetDf)
    if len(flaggedGrps) > 0:
        filled = True
        flaggedClassGrpDf = classGrpsDf[classGrpsDf['Name'].isin(flaggedGrps)]
        flaggedClassGrpDf.to_excel(writer,
                                       sheet_name='Flagged Class Groups',
                                       index=False)
        formatSheet('Flagged Class Groups', flaggedClassGrpDf)
    if len(flaggedVis) > 0:
        filled = True
        flaggedVisSheetDf = visDf.loc[flaggedVis].copy()
        for r, dt in enumerate(flaggedVisSheetDf['Date Last Used']):
            flaggedVisSheetDf.loc[flaggedVisSheetDf.index[r],
                                   'Date Last Used']=\
            datetime.datetime.strftime(dt, '%d/%m/%Y %H:%M:%S')
        flaggedVisSheetDf.to_excel(writer, sheet_name='Flagged Visualisations',
                                   index=False)
        formatSheet('Flagged Visualisations', flaggedVisSheetDf)
    if len(redundantConChecks) > 0:
        filled = True
        redundantConCheckSheetDf = conCheckDf.loc[redundantConChecks].copy()
        redundantConCheckSheetDf.reset_index(inplace=True)
        redundantConCheckSheetDf = redundantConCheckSheetDf[['Name',
                                                             'Description',
                                                             'Date Created',
                                                             'Date Last Used']]
        for r, dt in enumerate(redundantConCheckSheetDf['Date Last Used']):
            redundantConCheckSheetDf.loc[redundantConCheckSheetDf.index[r],
                                   'Date Last Used']=\
            datetime.datetime.strftime(dt, '%d/%m/%Y %H:%M:%S')
        redundantConCheckSheetDf.to_excel(writer,
                                          sheet_name='Redundant Con Checks',
                                          index=False)
        formatSheet('Redundant Con Checks', redundantConCheckSheetDf)
    if len(redundantImports) > 0:
        filled = True
        redundantImportsSheetDf = importsDf.loc[redundantImports].copy()
        redundantImportsSheetDf.reset_index(inplace=True)
        for r, dt in enumerate(redundantImportsSheetDf['Date Last Used']):
            redundantImportsSheetDf.loc[redundantImportsSheetDf.index[r],
                                   'Date Last Used']=\
            datetime.datetime.strftime(dt, '%d/%m/%Y %H:%M:%S')
        redundantImportsSheetDf.to_excel(writer,
                                          sheet_name='Redundant Imports',
                                          index=False)
        formatSheet('Redundant Imports', redundantImportsSheetDf)
    if len(redundantExports) > 0:
        filled = True
        redundantExportsSheetDf = exportsDf.loc[redundantExports].copy()
        redundantExportsSheetDf.reset_index(inplace=True)
        for r, dt in enumerate(redundantExportsSheetDf['Date Last Used']):
            redundantExportsSheetDf.loc[redundantExportsSheetDf.index[r],
                                   'Date Last Used']=\
            datetime.datetime.strftime(dt, '%d/%m/%Y %H:%M:%S')
        redundantExportsSheetDf.to_excel(writer,
                                          sheet_name='Redundant Exports',
                                          index=False)
        formatSheet('Redundant Exports', redundantExportsSheetDf)
    if filled:
        print('Saving...')
        writer.save()
    else:
        print(curStatAct + ' has no redundancies!')

def runTask():
    global configReports
    global grpSrchDf
    global redundantCalcs
    global flaggedTasks
    global flaggedGrps
    global flaggedVis
    global redundantParams
    global redundantImports
    global redundantExports
    global redundantConChecks
    global curStatAct
    global usedGrps
    global usedGrpsExt
    global inpFile, outFile
    configReports = pd.DataFrame(columns=['Config Report',
                                          'Statistical Activity',
                                          'Mode', 'Date'])
    grpSrchDf = pd.DataFrame()
    (inpFile, outFile) = ct.setupFilepaths()
    unzipFiles()
    createOutFolder()
    for i, configRpt in enumerate(configReports['Config Report']):
        curStatAct = configReports.loc[i, 'Statistical Activity']
        print('Writing redundancy report for ' + curStatAct + '...')
        redundantCalcs = []
        flaggedTasks = []
        flaggedGrps = []
        flaggedVis = []
        redundantParams = []
        redundantImports = []
        redundantExports = []
        redundantConChecks = []
        usedGrpsExt = []
        usedGrps = []
        os.chdir(inpFile)
        os.chdir(configRpt)
        readCalcs()
        readVisualisations()
        readConChecks()
        readImports()
        readExports()
        if not calculationsDf.empty:
            readTasks()
            readTaskLns()
            readParams()
            flagTasks()
        if not visDf.empty:
            flagVisualisations()
        if not conCheckDf.empty:
            findRedundantConChecks()
        if not importsDf.empty:
            findRedundantImports()
        if not exportsDf.empty:
            findRedundantExports()
        if not calculationsDf.empty:
            findRedundantCalcs()
            findRedundantParams()
        readClassGrps()
        readExtDatasets()
        if not classGrpsDf.empty:
            if not calculationsDf.empty or not extDatasetDf.empty:
                findRedundantClassGrps()
            readExtClass()
            if not extClassDf.empty:
                searchClassGrpsDependencies()
            compileUsedClassGrps()
        write()

timeTol = datetime.datetime.now() - datetime.timedelta(days=365)
runTask()
os.chdir(inpFile)
ct.done()