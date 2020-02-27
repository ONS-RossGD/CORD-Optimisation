# -*- coding: utf-8 -*-
import glob, os
import CORDtools as ct
import pandas as pd
import numpy as np
import networkx as nx
import matplotlib.pyplot as plt
import itertools as it


def createProductDf(df):
    df = df.replace('', np.nan)
    lists = []
    for col in df.columns:
        lists.append(df.loc[:, col].dropna().tolist())
    prodDf = pd.DataFrame(list(it.product(*lists)), columns=df.columns)
    prodDf = prodDf.drop_duplicates()
    prodDf = prodDf.reset_index(drop=True)
    return prodDf

def unpackClassMaps():
    global classMaps
    print('Unpacking Classification Mappings...')
    for file in glob.glob('*.csv'):
        #print(file)
        if 'MpItm' in file:
            try:
                meta = pd.read_csv(file, header=None, nrows=2,
                                   encoding='unicode_escape')
                mapItm = meta.loc[1,0].split('Mapping: ')[1].strip()
                mapDf = pd.read_csv(file, skiprows=3,
                                    encoding='unicode_escape',
                                    converters={'Source Code': lambda x: \
                                                str(x),
                                                'Target Code': lambda x: \
                                                str(x)})
                for r in mapDf.index:
                    src = mapDf.loc[r, 'Source Code'].strip()
                    targ = mapDf.loc[r, 'Target Code'].strip()
                    classMaps.loc[r, mapItm] = src+':'+targ
            except Exception as e:
                ct.error(e)
    print('Classification Mappings Unpacked!')

def unpackClassGroups():
    global classGrps
    print('Unpacking Classification Groups...')
    for file in glob.glob('*.csv'):
        if 'Grp' in file and 'Grps' not in file:
            try:
                meta = pd.read_csv(file, header=None, nrows=2,
                                   encoding='unicode_escape')
                info = meta.loc[1,0].split('Classification: ')
                info = info[1].split('Group: ')
                clas = info[0].strip()
                if clas == 'ZZZZ_SIC 2007 with EUROSTAT inds':
                    clas = 'Industry'
                grp = info[1].strip()
                header = clas +':'+ grp
                grpDf = pd.read_csv(file, skiprows=3,
                                    encoding='unicode_escape',
                                    converters={'Code': lambda x: str(x)})
                for r, code in enumerate(grpDf.loc[:, 'Code']):
                    classGrps.loc[r, header] = code.strip()
            except Exception as e:
                ct.error(e)
    print('Classification Groups Unpacked!')
    
def unpackPCLinks():
    global pcDf
    print('Unpacking Parent Child Links...')
    for file in glob.glob('*.csv'):
        if 'PCLnk' in file:
            try:
                meta = pd.read_csv(file, header=None, nrows=2,
                                   encoding='unicode_escape')
                clas = meta.loc[1,0].split('Classification: ')[1].strip()
                if clas == 'ZZZZ_SIC 2007 with EUROSTAT inds':
                    clas = 'Industry'
                pcTmp = pd.read_csv(file, skiprows=3,
                                    encoding='unicode_escape',
                                    converters={'Parent Code': lambda x: \
                                                str(x),
                                                'Child Code': lambda x: \
                                                str(x)})
                for r, parent in enumerate(pcTmp.loc[:, 'Parent Code']):
                    header = clas + ':' + parent
                    child = pcTmp.loc[r, 'Child Code']
                    if header not in pcDf.columns:
                        pcDf.loc[0, header] = child
                    if child not in pcDf[header].tolist():
                        pcDf.loc[pcDf[header].count(), header] = child
            except Exception as e:
                ct.error(e)
    for c, col in enumerate(pcDf.columns):
        for r, child in enumerate(pcDf[col]):
            clas = col.split(':')[0]
            child = str(child)
            header = clas + ':' + child
            if header in pcDf.columns:
                pcDf[col] = pcDf[col].dropna().append(pcDf[header],
                                                      ignore_index=True)
    print('Parent Child Links Unpacked!')
    
def unpackTasks():
    global taskGrps
    global taskDf
    global mappedIdxs
    print('Unpacking Tasks...')
    for file in glob.glob('*.csv'):
        try:
            taskDf = pd.read_csv(file, skiprows=6, header=None,
                                 encoding='unicode_escape')
            taskDf = taskDf.replace(np.nan, 'n/a')
            for r, task in enumerate(taskDf.loc[:, 2]):
                directMaps = []
                indirectMaps = []
                selCrit = taskDf.loc[r, 15]
                if selCrit != 'n/a':
                    for item in ct.splitCordString(selCrit, clas=False):
                        crit = item.split(' = ')[0].strip()
                        sel = item.split(' = ')[1].strip()
                        taskGrps.loc[crit, task.strip()] = sel
                #Don't add subtasks or delete calcs to the taskGrpDf
                if taskDf.loc[r, 13].strip() in {'SUB TASK', 'DELETE DATA'}:
                    continue
                taskGrps.loc['Type', task.strip()] = taskDf.loc[r, 13].strip()
                taskGrps.loc['Obj Dets', task.strip()] = \
                             taskDf.loc[r, 14].strip()
                taskGrps.loc['Order', task.strip()] = taskDf.loc[r, 0]
                taskGrps.loc['Target Dataset', task.strip()] = taskDf.loc[r, 9]
                if 'name = ' in taskDf.loc[r, 14]:
                    targData = taskDf.loc[r, 14].split('name = ')[1]\
                    .split('}')[0].strip()
                else:
                    targData = np.nan
                taskGrps.loc['Source Dataset', task.strip()] = targData
                dimMaps = taskDf.loc[r, 16]
                if 'Unmapped Dimension Mappings' not in taskGrps.index:
                    taskGrps.loc['Unmapped Dimension Mappings', task.strip()]\
                    = np.nan
                if 'Direct Dimension Mappings' not in taskGrps.index:
                    taskGrps.loc['Direct Dimension Mappings', task.strip()]\
                    = np.nan
                if 'Indirect Dimension Mappings' not in taskGrps.index:
                    taskGrps.loc['Indirect Dimension Mappings', task.strip()]\
                    = np.nan
                dims = ct.splitCordString(dimMaps, clas=False)
                for d, dim in enumerate(dims):
                    if '(Indirect)' in dim:
                        #print(dim)
                        #error('Indirect mapping found, it will be ignored!')
                        targCrit = dim.split('-> (Indirect)')[0].strip()
                        srcCrit = dim.split('-> (Indirect)')[1].strip()
                        indMap = dims[d+1]\
                                    .split('classification mapping = ')[1]\
                                    .split(' (from ')[0].strip()
                        del dims[d+1]                                    
                        #print(indMap)
                        
                        indirectMaps.append(srcCrit+':'+targCrit+':'+indMap)
                        taskGrps.loc['Indirect Dimension Mappings', 
                                     task.strip()] = ', '.join(indirectMaps)
                        continue
                    if '(Direct)' in dim:
                        crit = dim.split('-> (Direct) ')[1].strip()
                        origCrit = dim.split('-> (Direct) ')[0].strip()
                        if origCrit != crit:
                            directMaps.append(origCrit + ' = ' + crit)
                    if directMaps != []:
                        taskGrps.loc['Direct Dimension Mappings',
                                     task.strip()] = ', '.join(directMaps)
                    if '(Unmapped)' in dim:
                        crit = dim.split('-> (Unmapped) ')[1].strip()
                        sel = dim.split('-> (Unmapped) ')[0].strip() 
                        if pd.isnull(taskGrps.loc\
                                     ['Unmapped Dimension Mappings',
                                                  task.strip()]):
                                taskGrps.loc['Unmapped Dimension Mappings',
                                             task.strip()] = crit + ':' + sel
                        else:
                            taskGrps.loc['Unmapped Dimension Mappings',
                                             task.strip()] \
                                = taskGrps.loc['Unmapped Dimension Mappings',
                                               task.strip()]+'/'+crit+':'+sel
        except Exception as e:
            ct.error(e)
    print('Tasks Unpacked!')
    
def unpackClassifications():
    global classifications
    print('Unpacking Classifications...')
    for file in glob.glob('*.csv'):
        if 'Itms' in file:
            try:
                meta = pd.read_csv(file, header=None, nrows=2,
                                   encoding='unicode_escape')
                header = meta.loc[1,0].split('Classification: ')[1].strip()
                if header == 'ZZZZ_SIC 2007 with EUROSTAT inds':
                    header = 'Industry'
                if header == 'Adjustment Type':
                    header = 'Adjustment'
                classDf = pd.read_csv(file, skiprows=3,
                                    encoding='unicode_escape',
                                    converters={'Code': lambda x: str(x)})
                for r, code in enumerate(classDf.loc[:, 'Code']):
                    classifications.loc[r, header] = code.strip()
            except Exception as e:
                ct.error(e)
                
    classifications.loc[[0,1,2], 'Periodicity'] = ['A', 'Q', 'M']
    classifications.loc[range(13), 'Prices'] = ['CP', 'DEF', 'VM', 'CVM',
                        'PYP', 'CYP', 'IDEF', 'KQ', 'PYQ', 'CYQ', 'CRP', 'PRP',
                        'CVMRE']
    classifications.loc[range(13), 'Price'] = ['CP', 'DEF', 'VM', 'CVM',
                        'PYP', 'CYP', 'IDEF', 'KQ', 'PYQ', 'CYQ', 'CRP', 'PRP',
                        'CVMRE']
    print('Classifications Unpacked!')
    
def readTargSelCrit(calc):
    homePath = os.getcwd()
    os.chdir(procFile)
    preDf = pd.read_csv(calc+'.csv', engine='python')
    df = pd.read_csv(calc + '.csv', engine='python', converters={i: str \
                                for i in preDf.columns.tolist()})
    os.chdir(homePath)
    #print('read crit:\n', df)
    return df

def addNatAccs():
    global pcDf
    global classGrps
    homePath = os.getcwd()
    os.chdir(inputFile)
    dirs = filter(os.path.isdir, os.listdir())
    for aDir in dirs:
        if 'National Accounts_Config_Report' in aDir:
            os.chdir(aDir)
            for file in glob.glob('*.csv'):
                if 'PCLnk-CPA2008_235_Hierarchy' in file:
                    #meta = pd.read_csv(file, header=None, nrows=2,
                                   #encoding='unicode_escape')
                    #clas = meta.loc[1,0].split('Classification: ')[1].strip()
                    clas = 'CPA'
                    pcTmp = pd.read_csv(file, skiprows=3,
                                    encoding='unicode_escape',
                                    converters={'Parent Code': lambda x: \
                                                str(x),
                                                'Child Code': lambda x: \
                                                str(x)})
                    for r, parent in enumerate(pcTmp.loc[:, 'Parent Code']):
                        header = clas + ':' + parent
                        child = pcTmp.loc[r, 'Child Code']
                        if header not in pcDf.columns:
                            pcDf.loc[0, header] = child
                        if child not in pcDf[header].tolist():
                            pcDf.loc[pcDf[header].count(), header] = child
                if 'Grp' in file and 'CPA2008_235_Hierarchy' in file \
                and 'Grps' not in file:
                    try:
                        meta = pd.read_csv(file, header=None, nrows=2,
                                   encoding='unicode_escape')
                        info = meta.loc[1,0].split('Classification: ')
                        info = info[1].split('Group: ')
                        clas = 'CPA'
                        grp = info[1].strip()
                        header = clas +':'+ grp
                        grpDf = pd.read_csv(file, skiprows=3,
                                    encoding='unicode_escape',
                                    converters={'Code': lambda x: str(x)})
                        for r, code in enumerate(grpDf.loc[:, 'Code']):
                            classGrps.loc[r, header] = code.strip()
                    except Exception as e:
                        ct.error(e)
    for c, col in enumerate(pcDf.columns):
        for r, child in enumerate(pcDf[col]):
            clas = col.split(':')[0]
            child = str(child)
            header = clas + ':' + child
            if header in pcDf.columns:
                pcDf[col] = pcDf[col].dropna().append(pcDf[header],
                                                      ignore_index=True)
    os.chdir(homePath)

def saveSelCrit(calc, df):
    homePath = os.getcwd()
    os.chdir(procFile)
    df.to_csv(calc+'.csv', index=False)
    os.chdir(homePath)

def fillSelCritDfs():
    global inCrit
    global outCrit
    global undroppable
    global dropCols
    inCrit = pd.DataFrame(columns=taskGrps.columns.tolist())
    outCrit = pd.DataFrame(columns=taskGrps.columns.tolist())
    undroppable = pd.DataFrame(columns=taskGrps.columns.tolist())
    dropCols = ['Type', 'Order', 'Source Dataset', 'Target Dataset', 
                'Unmapped Dimension Mappings', 'Obj Dets',
                'Direct Dimension Mappings', 'Indirect Dimension Mappings']
    print('Filling selection criteria dataframes...')
    for calc in taskGrps.columns:
        #print(calc)
        inCrit[calc] = taskGrps[calc].drop(dropCols, axis=0)
        #Initially set output criteria = to input
        outCrit[calc] = taskGrps[calc].drop(dropCols, axis=0)
        calcType = taskGrps.loc['Type', calc]
        #Overwrite the criteria that are changing
        if calcType == 'FORMULA':
            if '{check {' in taskGrps.loc['Obj Dets', calc]:
                del taskGrps[calc]
                del inCrit[calc]
                del outCrit[calc]
                #taskGrps.drop(calc, axis=1, inplace=True)
                #error('Skipping consistency check')
                continue
            formStr = taskGrps.loc['Obj Dets', calc]\
                .split('formula = ')[1]
            formStr = formStr.split('}')[0].strip()
            newItems = taskGrps.loc['Obj Dets', calc]\
                .split('{items to calculate ')[1].split('}}')[0]
            newItems = newItems.replace('}', '')
            newItems = newItems.split('{')
            while '' in newItems:
                newItems.remove('')
            newCrits = []
            for item in newItems:
                item = str(item)
                crit = item.split(' = ')[0].strip()
                undroppable.loc[undroppable[calc].count(), calc] = crit
                sel = item.split(' = ')[1].strip()
                outCrit.loc[crit, calc] = sel
                newCrits.append(crit)
                formSelList = []
                for classItm in classifications.loc[:, crit]:
                    for char in ['+', '/', '-', '*', ',', ')', '(', ':']:
                        formStr = formStr.replace(char, ' ')
                    splitForm = formStr.split(' ')
                    #The below loop stops formula items such as 0 being 
                    #mistaken as classifications
                    for formEl in splitForm:
                        try:
                            float(formEl)
                            splitForm.remove(formEl)
                        except:
                            pass
                    if classItm in splitForm:
                        formSelList.append(classItm)
                formSel = ', '.join(formSelList)
                inCrit.loc[crit, calc] = formSel
            #outCrit.loc['Changes', calc] = ', '.join(newCrits)
        if calcType == 'AGGREGATE':
            parentType = taskGrps.loc['Obj Dets', calc]\
                .split('aggregate dimension = ')[1]
            parentType = parentType.split('}')[0].strip()
            parent = taskGrps.loc['Obj Dets', calc]\
                .split('top level specification = ')[1]
            parent = parent.split('}')[0].strip()
            try:
                inCrit.loc[parentType, calc] = ', '.join(pcDf[parentType+':'+
                      parent].dropna().tolist())
            except:
                ct.error('Couldnt find parent child group '+parentType+':'+
                      parent)
            outCrit.loc[parentType, calc] = parent
        if calcType == 'PRORATE':
            parentType = taskGrps.loc['Obj Dets', calc]\
                .split('hierarchy dimension = ')[1]
            parentType = parentType.split('}')[0].strip()
            parent = taskGrps.loc['Obj Dets', calc]\
                .split('top level specification = ')[1]
            parent = parent.split('}')[0].strip()
            measureType = taskGrps.loc['Obj Dets', calc]\
                .split('measure dimension = ')[1]
            measureType = measureType.split('}')[0].strip()
            measure = taskGrps.loc['Obj Dets', calc]\
                .split('prorate item = ')[1]
            measure = measure.split('}')[0].strip()
            inCrit.loc[parentType, calc] = parent
            try:
                outCrit.loc[parentType, calc] = ', '.join(pcDf[parentType+':'+
                      parent].dropna().tolist())
            except:
                ct.error('Couldnt find parent child group '+parentType+':'+
                      parent)
            inCrit.loc[measureType, calc] = measure
            outCrit.loc[measureType, calc] = measure
        if calcType == 'COPY BETWEEN DATASETS':
            if  pd.notnull(taskGrps.loc['Unmapped Dimension Mappings', calc]):
                dimMapStr = taskGrps.loc\
                ['Unmapped Dimension Mappings', calc].split('/')
                for dim in dimMapStr:
                    crit = dim.split(':')[0].strip()
                    sel = dim.split(':')[1].strip()
                    outCrit.loc[crit, calc] = sel

    for c, col in enumerate(inCrit.columns):
        for r, sel in enumerate(inCrit.loc[:, col]):
            sel = str(sel)
            splitSel = sel.split(',')
            for i, item in enumerate(splitSel):
                item = item.strip()
                if item[:5] == 'group':
                    item = item[6:]
                    found = False
                    for classCol in classGrps.columns:
                        if item not in classCol: continue
                        if inCrit.index[r] != classCol.split(':')[0]: continue
                        if item == classCol.split(':')[1]:
                            subitemsAsList = classGrps.loc[:, classCol]\
                                .dropna().tolist()
                            splitSel[i] = (', ').join(subitemsAsList)
                            found = True
                    if not found: ct.error('Group '+item+' not found!')
            inCrit.iloc[r, c] = ', '.join(splitSel)
    for c, col in enumerate(outCrit.columns):
        for r, sel in enumerate(outCrit.loc[:, col]):
            sel = str(sel)
            """if sel.strip() == '*':
                outCrit.loc[outCrit.index[r], col] = ', '.join(classifications\
                                   [outCrit.index[r]].dropna().tolist())
                continue"""
            splitSel = sel.split(',')
            for i, item in enumerate(splitSel):
                item = item.strip()
                if item[:5] == 'group':
                    item = item[6:]
                    for classCol in classGrps.columns:
                        if item not in classCol: continue
                        if outCrit.index[r] != classCol.split(':')[0]: continue
                        if item == classCol.split(':')[1]:
                            subitemsAsList = classGrps.loc[:, classCol]\
                                .dropna().tolist()
                            splitSel[i] = (', ').join(subitemsAsList)
            outCrit.iloc[r, c] = ', '.join(splitSel)
            
    for calc in outCrit.columns:
        dirDimMaps = taskGrps.loc['Direct Dimension Mappings', calc]
        indDimMaps = taskGrps.loc['Indirect Dimension Mappings', calc]
        #print(dirDimMaps)
        if pd.notnull(dirDimMaps):
            dirDims = dirDimMaps.split(', ')
            for dirDimMap in dirDims:
                origCrit = dirDimMap.split(' = ')[0].strip()
                crit = dirDimMap.split(' = ')[1].strip()
                #print(outCrit.loc[origCrit, calc], outCrit.loc[crit, calc])
                outCrit.loc[crit, calc] = outCrit.loc[origCrit, calc]
                outCrit.loc[origCrit, calc] = 'nan'
        if pd.notnull(indDimMaps):
            #print(indDimMaps)
            indDims = indDimMaps.split(', ')
            for indDim in indDims:
                srcCrit = indDim.split(':')[0].strip()
                targCrit = indDim.split(':')[1].strip()
                group = indDim.split(':')[2].strip()
                inVals = []
                outVals = []
                baseVals = classMaps[group].replace('nan', np.nan)
                for item in baseVals.dropna().tolist():
                    inVal = item.split(':')[1]
                    outVal = item.split(':')[0]
                    inVals.append(inVal)
                    outVals.append(outVal)
            #print('inVals:', inVals)
            #print('outVals:', outVals)
            outCrit.loc[targCrit, calc] = ', '.join(outVals)
            inCrit.loc[srcCrit, calc] = ', '.join(inVals)
                  
    print('Selection criteria dataframes filled!')
    
def permuteCriteria(critSeries, target=False):
    global unpermutedTargCriteria
    prelimCrit = pd.DataFrame(columns = critSeries.index)
    prelimCrit.loc[0] = critSeries.tolist()
    for r, cell in enumerate(prelimCrit.loc[0, :]):
        if ',' in str(cell):
            splitCell = cell.split(',')
            for c, val in enumerate(splitCell):
                prelimCrit.loc[c, prelimCrit.columns[r]] = str(val).strip()
    #Get rid of any duplicates within each column
    for col in prelimCrit.columns:
        for r in prelimCrit.index:
            if str(prelimCrit.loc[r, col])[:5] == 'group':
                prelimCrit.loc[r,col] = str(prelimCrit.loc[r, col])[6:]
        prelimCrit[col] = prelimCrit[col].drop_duplicates()
    prelimCrit.reset_index(inplace=True, drop=True)
    droppedCols = []
    prelimCrit = prelimCrit.replace('nan', np.nan)
    for col in prelimCrit.columns:
        if prelimCrit[col].isnull().all():
            droppedCols.append(col)
            prelimCrit = prelimCrit.drop(col, axis=1)
    if target:
        dfToSave = prelimCrit.copy()
        #print('Target Dataset: ', taskGrps.loc['Target Dataset',
        #              curAffectingCalc])
        dfToSave.loc[0, 'Target Dataset'] = taskGrps.loc['Target Dataset',
                      curAffectingCalc]
        #print('Saving targetSelCrit of', curAffectingCalc)
        homePath = os.getcwd()
        os.chdir(procFile)
        try:
            existingDf = pd.read_csv(curAffectingCalc+'.csv', engine='python')
            for col in existingDf.columns:
                for idx in existingDf.index:
                    if pd.isnull(existingDf.loc[idx, col]): continue
                    item = existingDf.loc[idx, col]
                    if item not in dfToSave[col].tolist():
                        dfToSave.loc[dfToSave[col].count()+1, col] = item
        except Exception as e:
            #error(e)
            pass
        os.chdir(homePath)
        #print(dfToSave)
        #print('Saving sel crit for', curAffectingCalc)
        #print(dfToSave)
        saveSelCrit(curAffectingCalc, dfToSave)
    selCritDf = createProductDf(prelimCrit)
    return selCritDf
    
def searchEffectedTasks(targDataset):
    global effectedDf
    global outCrit
    for calc in taskGrps.columns:
        if calc == 'Defined Selection Criteria': continue
        if taskGrps.loc['Order', calc] <= calcOrderNo: continue
        if targDataset not in {taskGrps.loc['Source Dataset', calc],
                               taskGrps.loc['Target Dataset', calc]}:
            continue
        inCritCopy = inCrit.copy()
        for col in targetCalcSelCrit.columns:
            if col not in inCrit.loc[inCrit[calc]!='nan', calc]:
                #print(col,'not in inCrit for', calc)
                inCritCopy.loc[col, calc] = ', '.join(targetCalcSelCrit[col]\
                              .drop_duplicates().tolist())
        curCalcSelCrit = permuteCriteria(inCritCopy[calc])
        refinedTargetSelCrit = targetCalcSelCrit.copy()
        #print(calc)
        #print('curCalcSelCrit:\n', curCalcSelCrit)
        #print('refinedTargSelCrit:\n', refinedTargetSelCrit)
        for col in refinedTargetSelCrit.columns:
            if '*' in refinedTargetSelCrit[col].tolist():
                refinedTargetSelCrit.drop(col, axis=1, inplace=True)
            if col in curCalcSelCrit.columns:
                if '*' in curCalcSelCrit[col].tolist():
                    curCalcSelCrit.drop(col, axis=1, inplace=True)
        try:
            matchedCrit = curCalcSelCrit.merge(refinedTargetSelCrit)
        except:
            matchedCrit = curCalcSelCrit
        nullMatch = False
        #print(curCalcSelCrit.dropna(axis=1, how='all').columns.tolist())
        #print(refinedTargetSelCrit.dropna(axis=1, how='all').columns.tolist())
        for col in curCalcSelCrit.dropna(axis=1, how='all').columns.tolist():
            if col not in refinedTargetSelCrit.dropna(axis=1, how='all').columns.tolist():
                nullMatch = True
        if nullMatch:
            matchedCrit = matchedCrit.iloc[0:0]
        if not matchedCrit.empty:
            #print(calc)
            #print(matchedCrit)
            r = len(effectedDf.index)
            if calc not in effectedDf['Effected Calc'].tolist():
                effectedDf.loc[r, 'Searched'] = False
            else:
                tempMatchedDf = effectedDf[effectedDf['Effected Calc'] == calc]
                #print('Matched Df\n', tempMatchedDf)
                if curAffectingCalc not in tempMatchedDf['Effected By']\
                                                .tolist():
                    effectedDf.loc[r, 'Searched'] = False
                else:
                    effectedDf.loc[r, 'Searched'] = True
            effectedDf.loc[r, 'Effected Calc'] = calc
            effectedDf.loc[r, 'Effected By'] = curAffectingCalc
            effectedDf.loc[r, 'Order'] = taskGrps.loc['Order', calc]
    
def userInput():
    global calcOrderNo
    global targetCalcSelCrit
    global curAffectingCalc
    calc = input('What is the Name of Calculation being changed?\n')
    if calc == 'end':
        return 'end'
    if calc == 'skip':
        return 'skip'
    while calc not in taskGrps.columns:
        ct.error('Calculation Name not found in task definition!')
        calc = input('What is the Name of Calculation being changed?\n')
    curAffectingCalc = calc
    calcOrderNo = taskGrps.loc['Order', calc]
    calcDf = pd.DataFrame()
    conf = 'N'
    while conf != 'Y':
        print('\nPlease confirm the selection criteria that is affected by',
              'the change (* for all):')
        for idx in outCrit[calc].index:
            if idx == 'Changes': continue
            if outCrit.loc[idx, calc] == 'nan': continue
            calcDf.loc[idx, calc] = input(idx + ' = ')
        print('\nPlease confirm these parameters are correct:')
        for idx in calcDf.index:
            print(idx, '=', calcDf.loc[idx, calc])
        conf = input('Enter Y to continue or N to redeifne the selection ' + 
                 'criteria: ')
    for crit in calcDf.index:
        if calcDf.loc[crit, calc] == '*':
            calcDf.loc[crit, calc] = outCrit.loc[crit, calc] 
                                    #Changed inCrit to outCrit
    targetCalcSelCrit = permuteCriteria(calcDf[calc], target=True)
    #print(targetCalcSelCrit)
    return 'default'

def fillImpactedDfs():
    global combinedDf
    os.chdir(procFile)
    combinedDf = pd.DataFrame()
    for file in glob.glob('*.csv'):
        df = pd.read_csv(file, engine='python')
        df['Target Dataset'] = df['Target Dataset'].replace(np.nan,
          df.loc[0, 'Target Dataset'])
        combinedDf = combinedDf.append(df)
    #combinedDf = combinedDf.replace('nan', np.nan)
    combinedDf.reset_index(drop=True, inplace=True)
    writer = pd.ExcelWriter('Dataset Impacts.xlsx', engine='xlsxwriter')
    for dataset in combinedDf['Target Dataset'].drop_duplicates()\
                    .dropna().tolist():
        tempDf = combinedDf[combinedDf['Target Dataset']==dataset].copy()
        tempDf.drop('Target Dataset', axis=1, inplace=True)
        impactedDf = pd.DataFrame(columns=tempDf.columns)
        for col in tempDf.columns:
            for item in tempDf[col].drop_duplicates().dropna().tolist():
                impactedDf.loc[impactedDf[col].count(), col] = item
            #print(datasetCritDf[dataset].dropna().tolist())
            if col not in datasetCritDf[dataset].dropna().tolist():
                impactedDf.drop(col, axis=1, inplace=True)
        #Future develpoment, identify item groups and '*' in impacted
        """for col in impactedDf.columns:
            for classGrp in classGrps.columns.tolist():
                Class = classGrp.split(':')[0]
                if Class != col: continue
                try:
                    if impactedDf[col].tolist() == classifications[col].tolist():
                        impactedDf.drop(col, axis=1, inplace=True)
                        impactedDf.loc[0, col] = '*'
                except:
                    pass
                if classGrps[classGrp].tolist() == impactedDf[col].tolist():
                    impactedDf.drop(col, axis=1, inplace=True)
                    impactedDf.loc[0, col] = classGrps[classGrp].tolist()"""
        os.chdir(outputFile)
        for char in ['[',']',':','*','?','\\','/']:
            dataset = dataset.replace(char, '')
        impactedDf.to_excel(writer, sheet_name=dataset, index=False)
    writer.save()
    #print(combinedDf)

def datasetCriteria():
    global datasetCritDf
    print('Creating dataset criteria dataframe...')
    for file in glob.glob('*.csv'):
        if 'Dataset_Definitions' not in file: continue
        try:
            prelimDatasetDf = pd.read_csv(file, engine='python', skiprows=3)
        except:
            ct.error('CORD has formatted Dataset Definitions badly. Please ' + 
                  'manually reformat Dataset_Definitions_rpt so that columns'+
                  ' align.')
        datasetCritDf = pd.DataFrame(columns=prelimDatasetDf['Name'].tolist())
        for r, dataset in enumerate(prelimDatasetDf['Name']):
            for col in prelimDatasetDf.columns.tolist():
                if 'Dimension' not in col: continue
                if prelimDatasetDf.loc[r, col].strip() == 'n/a': continue
                dim = prelimDatasetDf.loc[r, col].split('{name = ')[1]\
                    .split('}')[0].strip()
                datasetCritDf.loc[datasetCritDf[dataset].count(), dataset] = \
                dim
    print('Dataset criteria dataframe created!')

def delProcessing():
    print('Cleaning processing file...')
    homePath = os.getcwd()
    os.chdir(procFile)
    for file in glob.glob('*.csv'):
        os.remove(file)
    os.chdir(homePath)
    print('Cleaned sucessfully!')

def createGraph(effectedDf):
    G = nx.Graph()
    for idx in effectedDf.index:
        G.add_edge(effectedDf.loc[idx, 'Effected By'],
                   effectedDf.loc[idx, 'Effected Calc'])
    """pos = {'A2.2D Copy NATS from Input': (0,0),
           'A2.3C Shape Quarterly to RF': (1, 1),
           'A2.3D Forecast Quarterly R0': (1, -1),
           'A2.7A Calc NATS in Pounds £ms': (2, 0),
           'A2.14A Calc CAA_TOT': (3, 1),
           'A2.14L Calc OV_EXP_T': (3, -1),
           'A2.14O Calc OV_EXP_T VM': (4, 1),
           'A2.16I Calc EX 3.2.3N1': (4, -1),
           'A3.2B Copy RF to Adjustment data': (5, 0),
           'A3.2C Copy R1.1 data': (6, 0),
           'A3.3A Calc R1.2 data': (7, 0),
           'A4.3 Copy WIP to WIP': (8, 0)}"""
    pos = nx.spring_layout(G, k=0.3)
    nx.draw(G, pos, node_size=200, edge_color='r', alpha=0.4, font_size=12,
            with_labels=True)
    plt.show()
    
def fillDependancies():
    for file in glob.glob('*.csv'):
        if 'External_Datasets_Out_rpt_' not in file: continue
        tempDf = pd.read_csv(file, skiprows=3)
        dependanciesDf['Stat Act'] = tempDf['Stat Activity'].tolist()
        dependanciesDf['Mode'] = tempDf['Mode'].tolist()
        dependanciesDf['Task Name'] = tempDf['Name'].tolist()
        dependanciesDf['Effected Dataset'] = tempDf['Dataset'].tolist()
        for r, source in enumerate(tempDf['Type Specific Details']):
            dependanciesDf.loc[r, 'Source Dataset'] = source.split('name = ')\
                [1].split('}')[0].strip()
        """for r, selCritString in enumerate(tempDf['Selection Criteria']):
            selCritList = splitCordString(selCritString)
            for selCrit in selCritList:
                crit = selCrit.split(' = ')[0].strip()
                sel = selCrit.split(' = ')[1].strip()"""
        dependanciesDf['Selection Criteria'] = tempDf['Selection Criteria']

def checkChanges():
    global inCrit
    global outCrit
    for calc in inCrit.columns:
        changedList = []
        for idx in inCrit.index:
            if inCrit.loc[idx, calc] != outCrit.loc[idx, calc]:
                changedList.append(idx)
        outCrit.loc['Changes', calc] = ', '.join(changedList)

def checkDependencies():
    print('Checking dependancies...')
    os.chdir(outputFile)
    impactedDfList = []
    impactFile = pd.ExcelFile('Dataset Impacts.xlsx')
    for sheet in impactFile.sheet_names:
        df = impactFile.parse(sheet)
        df.loc[0, 'Dataset Name'] = sheet.strip()
        impactedDfList.append(df)
    overlap = list(set(impactFile.sheet_names) ^ \
            set(dependanciesDf['Effected Dataset']))
    if overlap != []:
        for impactedDf in overlap:
            for df in impactedDfList:
                if df.loc[0, 'Dataset Name'] != impactedDf: continue
                #df.drop('Dataset Name', axis=1, inplace=True)
                #Need to compare selection criteria between the impactedDf and
                #each of the rows in dependenciesDf that pick up data from the
                #impactedDf. If they share criteria, then that business area
                #will be effected.
                """critSeries
                for c in df.columns:
                    (', ').join(df[c].dropna().tolist())"""
    print('Dependencies checked!')

def mode():
    print('\n\n\n')
    print('Which mode would you like to run in?')
    print('1. Search via selection criteria')
    print('2. Search via calc')
    selMode = 99999
    while selMode not in {1,2}:
        selMode = input('Enter the number of your selected mode: ')
        selMode = int(selMode)
        if selMode not in {1,2}:
            ct.error('Unrecognised Mode!')
    return selMode
    
def searchLoop():
    global targetCalcSelCrit
    global calcOrderNo
    global curAffectingCalc
    global effectedDf
    while False in effectedDf['Searched'].tolist():
        #print(effectedDf)
        effectedDf = effectedDf.sort_values(by=['Order'])
        #print(effectedDf)
        searchedFalseIdxs = effectedDf.index[effectedDf['Searched']==False]
        c = min(searchedFalseIdxs.tolist())
        calc = effectedDf.loc[c, 'Effected Calc']
        print('Running search task for', calc,'...')
        curAffectingCalc = calc
        calcOrderNo = taskGrps.loc['Order', calc]
        #if outCrit.loc['Changes', calc] != np.nan:
        changed = outCrit.loc['Changes', calc].split(',')
        for cic, ci in enumerate(changed):
            changed[cic] = ci.strip()
        effectingCalc = effectedDf.loc[c, 'Effected By']
        effectingTargDf = readTargSelCrit(effectingCalc)
        targDataset = taskGrps.loc['Target Dataset', calc]
        effectingTargDf = effectingTargDf.drop('Target Dataset', axis=1)
        #print(effectingTargDf)
        effectedTargDf = outCrit.filter([calc], axis=1).T.copy()
        #print(effectedTargDf)
        effectedTargDf.drop('Changes', axis=1, inplace=True)
        effectedTargDf = effectedTargDf.replace('nan', np.nan)
        """relevantCols = datasetCritDf[targDataset].dropna().tolist()
        irrelevantCols = list(set(relevantCols) ^ 
                              set(effectedTargDf.columns.tolist()))
        print(irrelevantCols)
        effectedTargDf.drop(irrelevantCols, axis=1, inplace=True)"""
        effectedTargDf.dropna(axis=1, how='all', inplace=True)
        #effectedTargDf.dropna(how='all', axis=1, inplace=True)
        #print(effectedTargDf)
        for col in effectedTargDf.columns:
            if col in changed: continue
            replCol = effectingTargDf.loc[:, col].dropna().tolist()
            effectedTargDf.loc[calc, col] = ', '.join(replCol)
        effectedTargDf = effectedTargDf.T
        targetCalcSelCrit = permuteCriteria(effectedTargDf[calc], target=True)
        #saveSelCrit(curAffectingCalc, targetCalcSelCrit.copy())
        ######targetCalcSelCrit = permuteCriteria(outCrit[calc], target=True)
        #print('prev:\n', prevTargSelCrit)
        #print('new:\n', targetCalcSelCrit)
        #if calc == 'A3.2B Copy RF to Adjustment data':
        #print('Target DF\n', targetCalcSelCrit)
        searchEffectedTasks(targDataset)
        futureEffectedIdxs = effectedDf.index[effectedDf['Effected By']==calc]\
                                .tolist()
        for idx in futureEffectedIdxs:
            if idx <= c:
                futureEffectedIdxs.remove(idx)
        effectedDf.loc[futureEffectedIdxs, 'Searched'] = [False]\
                                                    *len(futureEffectedIdxs)
        effectedDf.loc[c, 'Searched'] = True
        effectedDf.drop_duplicates(subset=['Effected Calc', 'Effected By',
                                           'Searched'], inplace=True)
        effectedDf.reset_index(drop=True, inplace=True)
        #print(effectedDf)
        print(str(c+1) + '/' + str(len(effectedDf.index)))

def searchBySelCrit():
    global targetCalcSelCrit
    global calcOrderNo
    global curAffectingCalc
    calcOrderNo = 0
    #critList = [x for x in taskGrps.index.tolist() if x not in dropCols]
    for i, opt in enumerate(datasetCritDf.columns.tolist()):
        print(str(i)+'.', opt)
    dOpt = 999999
    print('Which dataset would you like to start the search from?')
    while dOpt > len(datasetCritDf.columns) or dOpt < 0:
        dOpt = input('Enter number corresponding to the desired dataset: ')
        dOpt = int(dOpt)
    dataset = datasetCritDf.columns[dOpt]
    taskGrps.loc['Target Dataset', 'Defined Selection Criteria'] = dataset
    print('Please enter the selections for each criteria.')
    print('If there is more than one selection, seperate it with a comma and',
          'a space.')
    print("If the criteria isn't required, leave it blank and hit enter.")
    critList = datasetCritDf[dataset].dropna().tolist()
    selCritDf = pd.DataFrame(index=critList)
    cont = 'N'
    while cont != 'Y':
        for  crit in critList:
            selCritDf.loc[crit, 'Defined Selection Criteria'] = \
            input(crit+' = ')
        print('Selection Criteria to search:')
        for idx in selCritDf.index.tolist():
            if selCritDf.loc[idx, 'Defined Selection Criteria'] != np.nan:
                print(idx, '=', selCritDf.loc[idx,
                                              'Defined Selection Criteria'])
        cont = input('Are these criteria correct? Enter Y to confirm: ')
    
    for idx in selCritDf.index:
        taskGrps.loc[idx, 'Defined Selection Criteria'] = selCritDf.loc[idx,
                    'Defined Selection Criteria']
    #selCritDf.loc['Target Dataset', 'Defined Selection Criteria'] = dataset
    #print(dataset)
    curAffectingCalc = 'Defined Selection Criteria'
    targetCalcSelCrit = permuteCriteria(selCritDf\
                                        ['Defined Selection Criteria'],
                                        target=True)
    searchEffectedTasks(dataset)
    searchLoop()

def searchByCalc():
    global targetCalcSelCrit
    global calcOrderNo
    global effectedDf
    global curAffectingCalc
    global roundCount
    uiVal = userInput()
    if uiVal == 'end': return
    searchEffectedTasks(taskGrps.loc['Target Dataset', curAffectingCalc])
    searchLoop()

def runTask():
    global targetCalcSelCrit
    global calcOrderNo
    global effectedDf
    global curAffectingCalc
    global roundCount
    dirs = list(filter(os.path.isdir, os.listdir()))
    roundCount = 0
    delProcessing()
    taskAct = ''
    for file in glob.glob('*.zip'):
        if os.path.splitext(file)[0] not in dirs:
            ct.unzipFiles(file)
            dirs.append(os.path.splitext(file)[0])
        if '_Task_Report_' in file:
            taskAct = file.split('_Task_Report_')[0]
    for aDir in dirs:
        if 'National Accounts_Config_Report' in aDir: continue
        if '_Config_Report_' in aDir and taskAct not in aDir:
            os.chdir(aDir)
            unpackClassifications()
            unpackClassGroups()
            unpackPCLinks()
            try:
                addNatAccs()
            except:
                ct.error('National Accounts Level config not found')
            os.chdir(inputFile)
        if taskAct+'_Task_Report_' in aDir:
            os.chdir(aDir)
            unpackTasks()
            os.chdir(inputFile)
        if taskAct+'_Config_Report_' in aDir:
            os.chdir(aDir)
            unpackClassMaps()
            fillDependancies()
            datasetCriteria()
            os.chdir(inputFile)
    fillSelCritDfs()
    checkChanges()
    
    os.chdir(outputFile)
    modeInt = mode()
    if modeInt == 1:
        searchBySelCrit()
    if modeInt == 2:
        searchByCalc()
    effectedDf.drop('Searched', axis=1, inplace=True)
    effectedDf.drop('Order', axis=1, inplace=True)
    print(effectedDf)
    try:
        effectedDf.to_excel('Effected Calcs.xlsx', index=False,
                        sheet_name='Effected Calcs')
    except:
        ct.error('Failed to save Effected Calcs!')
    fillImpactedDfs()
    createGraph(effectedDf)
    checkDependencies()
            
classifications = pd.DataFrame()
classGrps = pd.DataFrame()
classMaps = pd.DataFrame()
taskGrps = pd.DataFrame()
pcDf = pd.DataFrame()
datasetImpactDf = pd.DataFrame()
dependanciesDf = pd.DataFrame(columns=['Stat Act', 'Mode', 'Task Name',
                                       'Effected Dataset', 'Source Dataset',
                                       'Selection Criteria'])
effectedDf = pd.DataFrame(columns=['Effected Calc', 'Effected By', 'Searched'])
(inputFile, procFile, outputFile) = ct.setupFilepaths(proc=True)
runTask()