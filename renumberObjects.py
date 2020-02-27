# -*- coding: utf-8 -*-
"""
Script to create a specification to renumber the codes of Stat Act Objects
(Task, calcs, consistency checks).
"""

import pandas as pd, numpy as np, CORDtools as ct
import glob, os

def createConcatTaskTree():
    global concatTaskTree
    concatTaskTree = taskTree.copy()
    for task in concatTaskTree.columns:
        for obj in concatTaskTree[task].dropna():
            if obj in taskTree.columns:
                for subObj in taskTree[obj].dropna():
                    if subObj not in concatTaskTree[task]:
                        concatTaskTree.loc[concatTaskTree[task].count(), task]\
                        =subObj

def readTaskLns():
    global taskTree
    global objectDf
    children = []
    taskTree = pd.DataFrame()
    objectDf = pd.DataFrame()
    for file in glob.glob('*.csv'):
        if 'TskLn-' in file:
            prelimDf = pd.read_csv(file, sep='|', header=None,
                                   encoding='unicode_escape')
            parent = prelimDf.loc[1, 0].split('Task: ')[1]\
                     .replace(' (Not Parallel)', '').replace(' (Parallel)', '')
            taskLnDf = pd.read_csv(file, skiprows=3, encoding='unicode_escape')
            taskTree[parent] = np.nan
            for r, child in enumerate(taskLnDf['Name']):
                if child not in objectDf.index:
                    objectDf.loc[child, 'Type'] = taskLnDf.loc[r, 'Type']
                children.append(child)
                taskTree.loc[taskTree[parent].count(), parent] = child

def defineTopLevel():
    global topLevels
    done = False
    print('Please define which tasks are top level for the Stat Act',
          curStatAct)
    print('When you have finished defining the top level tasks, enter "DONE".')
    topLevels = []
    while done == False:
        taskName = input('Top Level Task: ')
        if taskName in {'DONE', 'done',''}:
            done = True
        else:
            if taskName in taskTree.columns.tolist():
                try: 
                    float(taskName.split(' ')[0])
                    topLevels.append(taskName)
                except:
                    ct.error('Please make sure the top level task has a valid'+
                          ' numeric code prefix. Unable to use this task as'+
                          ' top level.', warning=True)
            else:
                ct.error('No task with the name "'+taskName+'" exists.',
                      warning=True)

def stripExistingCode(name):
    hasCode = False
    prefix = name.split(' ')[0]
    try:
        tempPre = prefix
        float(tempPre)
        hasCode = True
    except:
        pass
    if '.' in prefix:
        hasCode = True
    if hasCode:
        name = (' ').join(name.split(' ')[1:]).strip()
    return name
        
def addCodes():
    global renumberDf
    for oldName in codesDf.index:
        newName = codesDf.loc[oldName, 'Code']+' '+stripExistingCode(oldName)
        tempDf = pd.DataFrame({'Old Name':[oldName],'New Name':[newName]})
        renumberDf = renumberDf.append(tempDf, ignore_index=True)

def renumber(topLevel):
    global codesDf
    codesDf.loc[topLevel, 'Code'] = str(topLevel.split(' ')[0])
    taskList = []
    for i, obj1 in enumerate(taskTree[topLevel].dropna()):
        codesDf.loc[obj1, 'Code'] = codesDf.loc[topLevel, 'Code']+'.'\
                                    +str(i)
        if obj1 in taskTree.columns.tolist():
            taskList.append(obj1)
    for curTask in taskList:
        for i, obj in enumerate(taskTree[curTask].dropna()):
            codesDf.loc[obj, 'Code'] = codesDf.loc[curTask, 'Code']\
                                        + '.' + str(i)
            if obj in taskTree.columns.tolist():
                taskList.append(obj)

def assignObjType():
    global renumberDf
    for r, obj in enumerate(renumberDf['Old Name']):
        try:
            renumberDf.loc[r, 'Type'] = objectDf.loc[obj, 'Type']
        except:
            renumberDf.loc[r, 'Type'] = 'TASK'
    renumberDf = renumberDf[['Type','Old Name','New Name']]

def createFormat(df, style):
    """Creates the xlsxwriter format for a given dataframe and returns it as a
    string.
    """
    cols = []
    for col in df.columns.tolist():
        cols.append({'header':col})
    form = {'columns':cols, 'style': style}
    return form

def formatSheet(sheet, df):
    wb = writer.book
    ws = writer.sheets[sheet]
    reviewForm = wb.add_format({'bg_color': '#FFC7CE',
                                'font_color': '#9C0006'})
    sameForm = wb.add_format({'bg_color': '#C6EFCE',
                              'font_color': '#006100'})
    if sheet == 'Specification':
        style = 'Table Style Medium 5'
        ws.add_table(0,0,len(df.index),2, createFormat(df, style))
        ws.conditional_format(1,2,len(df.index),2,{'type':     'formula',
                                              'criteria': '=LEN($C2)>32',
                                              'format':   reviewForm
                                              })
        ws.conditional_format(1,0,len(df.index),2,{'type':     'formula',
                                              'criteria': '$B2=$C2',
                                              'format':   sameForm
                                              })
        ws.set_column('A:A', 24)
        ws.set_column('B:B', 30)
        ws.set_column('C:C', 30)

def write():
    global writer
    os.chdir(outFolder)
    print('Writing spec for '+curStatAct+'...')
    writer = pd.ExcelWriter(curStatAct+' Renumbered.xlsx',
                            engine='xlsxwriter')
    renumberDf.to_excel(writer, sheet_name='Specification',
                         index=False)
    formatSheet('Specification', renumberDf)
    writer.save()

def runTask():
    global configReports
    global curStatAct
    global taskRenumberingDf
    global renumberDf
    global codesDf
    global inpFol, outFolder
    codesDf = pd.DataFrame(columns=['Code'])
    (inpFol, outFol) = ct.setupFilepaths()
    configReports = ct.unzipConfigReports(inpFol)
    outFolder = ct.createOutFolder(outFol)
    for i, configRpt in enumerate(configReports['Config Report']):
        renumberDf = pd.DataFrame(columns=['Old Name','New Name'])
        os.chdir(inpFol)
        os.chdir(configRpt)
        curStatAct = configReports.loc[i, 'Statistical Activity']
        readTaskLns()
        print()
        defineTopLevel()
        print('Renumbering objects for ' + curStatAct + '...')
        for topLevel in topLevels:
            renumber(topLevel)
            addCodes()
        assignObjType()
        write()
        
runTask()
os.chdir(inpFol)
ct.done()