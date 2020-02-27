# -*- coding: utf-8 -*-
"""
Created on Fri Jan 24 15:29:04 2020

@author: gregor1
"""

import pandas as pd
import glob, os, datetime
import CORDtools as ct

def updateMasters():
    global outMasterDf
    global inMasterDf
    if not extDataOutDf.empty:
        for r in extDataOutDf.index.tolist():
            topLine = len(outMasterDf.index.tolist())
            outMasterDf.loc[topLine, 'Source'] = curStatAct
            outMasterDf.loc[topLine, 'Target'] = extDataOutDf.loc[r,
                                                            'Stat Activity']
            outMasterDf.loc[topLine, 'Copy Calc'] = extDataOutDf.loc[r, 'Name']
            selCritList = ct.splitCordString(extDataOutDf.loc[r,
                                                         'Selection Criteria'])
            dimMapList = ct.splitCordString(extDataOutDf.loc[r,
                                                         'Dimension Mappings'])
            selCritDf = pd.DataFrame()
            for c1, selCrit in enumerate(selCritList):
                outMasterDf.loc[topLine+c1, 'Data Sent'] = selCrit
                splitSelCrit = selCrit.split(' = ')
                selCritDf.loc[splitSelCrit[0], 'Value'] = splitSelCrit[1]
            for c2, dimMap in enumerate(dimMapList):
                dimType = dimMap.split('-> ')[1].split('(')[1].split(')')[0]
                splitDim = dimMap.split('->')
                left = splitDim[1].split(') ')[1].strip()
                if dimType == 'Direct':
                    right = selCritDf.loc[splitDim[0].strip(), 'Value']
                if dimType == 'Unmapped':
                    right = dimMap.split('-> ')[0]
                if dimType == 'Indirect':
                    splitLeft = left.split(' classification mapping = ')
                    left = splitLeft[0]
                    right = selCritDf.loc[splitDim[0].strip(), 'Value']
                    right += ' (Mapped using ' + splitLeft[1] + ')'
                outMasterDf.loc[topLine+c2, 'Data Recieved'] = left+' = '+right
    if not extDataInDf.empty:
        for r in extDataInDf.index.tolist():
            topLine = len(inMasterDf.index.tolist())
            longSrcStr = extDataInDf.loc[r, 'Type Specific Details']
            inMasterDf.loc[topLine, 'Source'] = longSrcStr.split('(from ')[1]\
                                                .split(')')[0]
            inMasterDf.loc[topLine, 'Target'] = curStatAct
            inMasterDf.loc[topLine, 'Copy Calc'] = extDataInDf.loc[r, 'Name']
            selCritList = ct.splitCordString(extDataInDf.loc[r,
                                                         'Selection Criteria'])
            dimMapList = ct.splitCordString(extDataInDf.loc[r,
                                                         'Dimension Mappings'])
            selCritDf = pd.DataFrame()
            for c1, selCrit in enumerate(selCritList):
                inMasterDf.loc[topLine+c1, 'Data Sent'] = selCrit
                splitSelCrit = selCrit.split(' = ')
                selCritDf.loc[splitSelCrit[0], 'Value'] = splitSelCrit[1]
            for c2, dimMap in enumerate(dimMapList):
                dimType = dimMap.split('-> ')[1].split('(')[1].split(')')[0]
                splitDim = dimMap.split('->')
                left = splitDim[1].split(') ')[1].strip()
                if dimType == 'Direct':
                    right = selCritDf.loc[splitDim[0].strip(), 'Value']
                if dimType == 'Unmapped':
                    right = dimMap.split('-> ')[0]
                if dimType == 'Indirect':
                    splitLeft = left.split(' classification mapping = ')
                    left = splitLeft[0]
                    right = selCritDf.loc[splitDim[0].strip(), 'Value']
                    right += ' (Mapped using ' + splitLeft[1] + ')'
                inMasterDf.loc[topLine+c2, 'Data Recieved'] = left+' = '+right

def readExtDatasets():
    global extDataOutDf
    global extDataInDf
    extDataOutDf = pd.DataFrame()
    extDataInDf = pd.DataFrame()
    for file in glob.glob('*.csv'):
        if 'External_Datasets_Out_rpt_' in file:
            extDataOutDf = pd.read_csv(file, skiprows=3,
                                       encoding='unicode_escape')
        if 'External_Datasets_In_rpt_' in file:
            extDataInDf = pd.read_csv(file, skiprows=3,
                                      encoding='unicode_escape')

def formatSheet(sheet, df):
    wb = writer.book
    ws = writer.sheets[sheet]
    boldForm = wb.add_format({'bold':1})
    boldForm.set_top(2)
    if sheet == 'External Dependencies':
        style = 'Table Style Medium 5'
        ws.add_table(0,0,len(df.index),4, ct.createFormat(df, style))
        ws.conditional_format(1,0,len(df.index),4,{'type': 'formula',
                              'criteria': '$E2<>$E1',
                              'format': boldForm})
        ws.set_column('A:A', 30)
        ws.set_column('B:B', 30)
        ws.set_column('C:C', 56)
        ws.set_column('D:D', 56)
        ws.set_column('E:E', 30)
    if sheet == 'Internal Dependencies':
        style = 'Table Style Medium 5'
        ws.add_table(0,0,len(df.index),4, ct.createFormat(df, style))
        ws.conditional_format(1,0,len(df.index),4,{'type': 'formula',
                              'criteria': '$E2<>$E1',
                              'format': boldForm})
        ws.set_column('A:A', 30)
        ws.set_column('B:B', 30)
        ws.set_column('C:C', 56)
        ws.set_column('D:D', 56)
        ws.set_column('E:E', 30)

def write():
    global writer
    os.chdir(outFol)
    timeStr = datetime.datetime.now().strftime("%d-%m-%Y (%H;%M;%S)")
    writer = pd.ExcelWriter('Dependencies Report '+timeStr+'.xlsx',
                            engine='xlsxwriter')
    outMasterDf.to_excel(writer, sheet_name='External Dependencies',
                         index=False)
    inMasterDf.to_excel(writer, sheet_name='Internal Dependencies',
                         index=False)
    formatSheet('External Dependencies', outMasterDf)
    formatSheet('Internal Dependencies', inMasterDf)
    writer.save()

def runTask():
    global outMasterDf
    global inMasterDf
    global configReports
    global curStatAct
    global inpFol, outFol
    outMasterDf = pd.DataFrame(columns=['Source','Target', 'Data Sent', 
                                     'Data Recieved', 'Copy Calc'])
    inMasterDf = pd.DataFrame(columns=['Source','Target', 'Data Sent', 
                                     'Data Recieved', 'Copy Calc'])
    (inpFol, outFol) = ct.setupFilepaths()
    configReports = ct.unzipConfigReports(inpFol)
    for i, configRpt in enumerate(configReports['Config Report']):
        curStatAct = configReports.loc[i, 'Statistical Activity']
        os.chdir(inpFol)
        os.chdir(configRpt)
        print('Reading external dependencies for ' + curStatAct + '...')
        readExtDatasets()
        updateMasters()
    outMasterDf['Source'].fillna(method='ffill', inplace=True)
    outMasterDf['Target'].fillna(method='ffill', inplace=True)
    outMasterDf['Copy Calc'].fillna(method='ffill', inplace=True)
    inMasterDf['Source'].fillna(method='ffill', inplace=True)
    inMasterDf['Target'].fillna(method='ffill', inplace=True)
    inMasterDf['Copy Calc'].fillna(method='ffill', inplace=True)
    write()
    
runTask()
os.chdir(inpFol)
ct.done()