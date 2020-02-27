# -*- coding: utf-8 -*-
"""Script to compare CORD dataset extracts.

The extracts should follow the format:
    - Page: Periodicity *ONLY*
    - Column: Date *ONLY*
    - Row: All other dimensions (can be in any order)
    - Download Type: CSV

The downloaded filenames of the extracts should not be changed since the script
can automatically deduce which extract is the Pre/Post from the default file
names. If it is necessary to change the filename, the numeric string at the end
of the filename (eg "..._123456_123456.csv") must remain intact. It is also 
important not to open and save the file in Excel before passing it to the
script since this removes any leading and trailing zeros.

All CSV extracts should be placed in the INPUT folder. An INPUT folder will be
created on the first run if it doesn't already exist.

There are a number of User Options at the bottom of the script where the output
options can be adjusted. By default, a formatted xlsx spreadsheet will be
generated containing the Pre and Post data alongside the Difference dataset.
The Difference dataset is the Post dataset - the Pre dataset.

Author: Ross Gregory-Davies : gregor1
"""
import pandas as pd
import os, sys, glob, datetime, warnings, time
import numpy as np
import matplotlib.pyplot as plt
import CORDtools as ct

def prelimReadCsv(file):
    """Reads the csv as a single column dataframe in order to gather info
    about the files structure and metadata.
    
    Returns a dataframe (dimDf) with the line a dataset starts as index, the
    dimensions it contains in column 0, number of lines to be read in the
    dataset in column 1, and periodicity in column 2."""
    df = pd.read_csv(file, sep='|', skip_blank_lines=False, 
                     keep_default_na=False, header=None, dtype=str)
    dimDf = df[df[0].str.contains(',Date')].copy()
    perDf = df[df[0].str.contains(', Periodicity:')].copy()
    if dimDf.empty:
        ct.error('File: "'+ file +'" is not in the correct CORD csv format.')
        sys.exit()
    blanksDf = df[df[0]==''].copy()
    badFile = False
    if blanksDf.empty:
        ct.error(file + ' has been previously opened and saved in Excel. '+
              'Leading zeros will be attempted to be added back in but '+
              'trailing zeros have been perminantly lost. For '+
              'best results, please avoid modifying the file the future.',
              warning=True)
        badFile = True
        for r, line in enumerate(df[0]):
            if line == len(line)*line[0]:
                blanksDf.loc[r, 0] = ''
        blanksDf.loc[len(df.index.tolist()), 0] = ''
    for c, r in enumerate(dimDf.index):
        finish = blanksDf.index[blanksDf.index>=int(r)].tolist()
        if finish == []:
            finish = [df.index.tolist()[-1]]
            #-2 is so that we read past the dimension line and date header
        dimDf.loc[r, 1] = finish[0]-r-2
        if perDf.empty:
            try:
                dimDf.loc[r, 2] = df[df[0].str.match('Periodicity,')]\
                .values[0][0].split(',')[1].strip().strip('"')
            except:
                ct.error('File format error regarding periodicity.')
                sys.exit()
        else:
            dimDf.loc[r, 2] = perDf.iloc[c, 0].split(', Periodicity:')[1]\
                            .split(',')[0].strip()
        headers = df.loc[r+1, 0].replace('"','').split(',')
        while '' in headers: headers.remove('')
        dimDf.loc[r, 3] = ','.join(headers)
        dimDf.loc[r, 4] = badFile
    return dimDf
        
def genDtypes(dims, headers):
    """Generates the header names and their respective data types (str for all
    dimension columns and floats for everything else). Used when reading a csv.
    Returns:
        - names: The header names as a list
        - dtypes: The data types per column as a list
    """
    names = dims[:-1]+headers
    dList = []
    for dim in dims[:-1]:
        dList.append('str')
    for head in headers:
        dList.append('float64')
    dtypes = dict(zip(names,dList))
    return names, dtypes


def readCORDcsv(file, leadingZeros=True):
    """Reads a CORD csv, so long as it was extracted with ONLY date in Columns
    and ONLY periodicity (if available) in Pages.
    Returns a list, dfs, that contains the read dataframe and it's periodicity.
    """
    dfs = []
    dimDf = prelimReadCsv(file)
    for r in range(len(dimDf.index)):
        #Pull the metadata elements identified by the preliminary csv read
        skip = dimDf.index[r]+2
        dims = dimDf.iloc[r, 0].split(',')
        headers = dimDf.iloc[r, 3].split(',')
        while '' in dims: dims.remove('')
        noOfDims = len(dims)-1
        names, dtypes = genDtypes(dims, headers)
        #Read the csv with as many defined elements as possible
        df = pd.read_csv(file, skiprows=skip, keep_default_na=False,
                         na_values=['.','NULL',''], index_col=False,
                         skip_blank_lines=False, low_memory=False,
                         header=None, names=names, dtype=dtypes,
                         nrows=dimDf.iloc[r, 1], usecols=range(len(names)))
        #Rename the dimension columns to their proper titles
        for d in range(noOfDims):
            #Forward fill the index columns ready for multiindexing
            df[dims[d]].fillna(method='ffill', inplace=True)
        #Drop any leftover unnamed columns
        df.drop(df.columns[df.columns.str.contains('Unnamed: ')].tolist(),
                           axis=1, inplace=True)
        if dimDf.iloc[r, 4] == True:
            if leadingZeros==True:
                for dim in dims[:-1]:
                    for i, val in enumerate(df[dim]):
                        try:
                            splitVal = val.split('.')[0]
                            valAsFloat = float(splitVal)
                            if valAsFloat < 10:
                                df.loc[i, dim] = '0'+df.loc[i, dim]
                                ct.error('Adding leading zero to create '+
                                      df.loc[i, dim], warning=True)
                        except:
                            pass
        #Set the dimension columns as a multiindex
        df.set_index(dims[:-1], inplace=True)
        #Re-order the df to make the years follow a logical order again
        df = df[headers]
        dfs.append((df, dimDf.iloc[r, 2]))
    return dfs

def readPreChange(lZeros):
    """Read the Pre-Change files identified by configureFiles. Add their dfs
    (returned by readCORDcsv) to the filesDf Dataframe.
    """
    global filesDf
    os.chdir(inpFol)
    for i, file in enumerate(preFiles):
        print('Reading Pre-Change file for', file[:-18]+'...')
        r = filesDf['Pre File Name'].count()
        dfs = readCORDcsv(file, lZeros)
        filesDf.loc[r, 'Pre File Name'] = file
        filesDf.loc[r, 'Post File Name'] = postFiles[i]
        pers = []
        for df, p in dfs:
            pers.append(p)
        filesDf.loc[r, 'Periodicities'] = ','.join(pers)
        filesDf.loc[r, 'Pre'] = dfs
     
    
def readPostChange(lZeros):
    """Read the Post-Change files identified by configureFiles. Add their dfs
    (returned by readCORDcsv) to the filesDf Dataframe. Also ensures that the
    Post-Change file matches the format of the Pre-Change file which has
    already been read.
    """
    global filesDf
    os.chdir(inpFol)
    for idx, file in enumerate(filesDf['Post File Name'].tolist()):
        print('Reading Post-Change file for', file[:-18]+'...')
        dfs = readCORDcsv(file, lZeros)
        pers = []
        for i, (df, p) in enumerate(dfs):
            if df.index.names != filesDf.loc[idx, 'Pre'][0][0].index.names:
                ct.error('Index Format in Post file is different to that of Pre.'
                      ' Reordering Post Index to match Pre.', warning=True)
                try:
                    dfs[i] = (df.reorder_levels(filesDf.loc[idx, 'Pre'][0][0]\
                               .index.names), p)
                except:
                    ct.error('Unable to reorder Post Index to match Pre, likely'
                          ' because the dimensions are different. Please fix'
                          ' this file before continuing!...')
                    sys.exit()
            pers.append(p)
        prePers = filesDf.loc[idx, 'Periodicities'].split(',')
        if len(pers) != len(prePers):
            ct.error('File "'+file+'" contains a different number of pages '+
                  '(Periodicities) in Post-Change to that in Pre-Change. ' +
                  'Periodicities that arent shared will not be compared.',
                  warning=True)
            dropIdxs = []
            diff = list(set(pers)^set(prePers))
            for i, (j, k) in enumerate(dfs):
                if k in diff:
                    dropIdxs.append(i)
            for i in dropIdxs:
                ct.error('Dropping periodicity '+pers[i], warning=True)
                del dfs[i]
        filesDf.loc[idx, 'Post'] = dfs

def configureNans(preDf,postDf,compDf,method='optimal'):
    """Fills all blank values with 'Not In Pre', 'Not In Post', or '.'. There
    are 3 methods of achieving this.
        - fast: Holds the entire dataframe in memory and calculates all NaN
                Values in one go. This is the fastest method but also is
                heavily memory dependant.
                Note: Current version has issues and should not be used.
        - optimal: Calculates all NaN values column by column. A lot less
                memory dependant than 'fast' but not much slower. This is the
                default option.
        - lowmem: Calculates all NaN values a single value at a time. Much 
                slower than the other 2 methods and should only be used on
                dataframes that have so many rows that the optimal method runs
                out of memory.
    """
    if method not in {'optimal','fast','lowmem', 'none'}:
        ct.error('Unknown NaN configuration method.')
        sys.exit()
    colsNotInPre = list(set(compDf.columns.values).symmetric_difference(\
                         preDf.columns.values))
    colsNotInPost = list(set(compDf.columns.values).symmetric_difference(\
                         postDf.columns.values))
    rowsNotInPre = list(set(compDf.index.values).symmetric_difference(preDf\
                         .index.values))
    rowsNotInPost = list(set(compDf.index.values).symmetric_difference(postDf\
                         .index.values))
    compDf.loc[:, colsNotInPre] = 'Col Not In Pre'
    compDf.loc[:, colsNotInPost] = 'Col Not In Post'
    compDf.loc[rowsNotInPre, :] = 'Row Not In Pre'
    compDf.loc[rowsNotInPost, :] = 'Row Not In Post'
    if method == 'lowmem':
        nanLocs = np.where(pd.isnull(compDf))
        for [r,c] in zip(*nanLocs):
            row = compDf.index[r]
            col = compDf.columns[c]
            if pd.isnull(postDf.loc[row,col]) and \
            pd.isnull(preDf.loc[row,col]):
                compDf.iloc[r,c] = '.'
            if pd.isnull(postDf.loc[row,col]) and \
            not pd.isnull(preDf.loc[row,col]):
                compDf.iloc[r,c] = 'Not In Post'
            if not pd.isnull(postDf.loc[row,col]) and \
            pd.isnull(preDf.loc[row,col]):
                compDf.iloc[r,c] = 'Not In Pre'
    if method == 'fast':
        """There is an issue with this method that is as yet unresolved. For 
        the time being, please use optimal method."""
        postIndDf = pd.DataFrame(columns=['Rows', 'Columns'])
        preIndDf = pd.DataFrame(columns=['Rows', 'Columns'])
        #Pick up the iloc locations for nan values in pre and post)
        preNans = np.where(pd.isnull(preDf))
        postNans = np.where(pd.isnull(postDf))
        #Convert the iloc values back to named indices and columns
        preRows = list(preDf.index[list(preNans[0])].values)
        preCols = list(preDf.columns[list(preNans[1])].values)
        postRows = list(postDf.index[list(postNans[0])].values)
        postCols = list(postDf.columns[list(postNans[1])].values)
        #Add the named values to a dataframe
        preIndDf['Rows'] = preRows
        preIndDf['Columns'] = preCols
        postIndDf['Rows'] = postRows
        postIndDf['Columns'] = postCols
        #Concat the two dataframes
        mergeIndDf = pd.concat([preIndDf, postIndDf])
        #Drop the duplicate cell locations
        mergeIndDf.drop_duplicates(subset=['Rows', 'Columns'],
                  inplace=True)
        #Get the rows and cols of mergeIndDf that match the listed rows and
        #cols that have already been marked for removal earlier on
        if mergeIndDf.empty:
            dropCols, dropRows = [], []
        else:
            dropCols = mergeIndDf['Columns'].str.match('|'.join(colsNotInPost+
                                 colsNotInPre)).dropna().tolist()
            dropRows = mergeIndDf['Rows'].str.match('|'.join(map(str,
                                 rowsNotInPost+rowsNotInPre))).dropna()\
                                 .tolist()
        #Get the indices of the matchin column/row items as a list
        dropColInd = mergeIndDf.index[dropCols].tolist()
        dropRowInd = mergeIndDf.index[dropRows].tolist()
        #Drop those rows and cols from the lists in the mergedIndDf
        mergeIndDf.drop(list(dict.fromkeys(dropColInd+dropRowInd)), axis=0,
                        inplace=True)
        #Take a copy of post and pre of only the merged rows and cols
        preCopy = preDf.loc[mergeIndDf['Rows'].tolist(),
                                  mergeIndDf['Columns'].tolist()].copy()
        postCopy = postDf.loc[mergeIndDf['Rows'].tolist(),
                                    mergeIndDf['Columns'].tolist()].copy()
        #Find the iloc locations of the non-nan values in both post and pre
        notInPostCopy = np.where(pd.notnull(preCopy))
        notInPreCopy = np.where(pd.notnull(postCopy))
        #Convert the ilocs back to named indices
        notInPostRows = list(preCopy.index[list(notInPostCopy[0])].values)
        notInPostCols = list(preCopy.columns[list(notInPostCopy[1])].values)
        notInPreRows = list(postCopy.index[list(notInPreCopy[0])].values)
        notInPreCols = list(postCopy.columns[list(notInPreCopy[1])].values)
        #Mass mark those cells as not in pre/post
        compDf.loc[notInPostRows, notInPostCols] = 'Not In Post'
        compDf.loc[notInPreRows, notInPreCols] = 'Not In Pre'
        compDf.replace(np.nan, '.', inplace=True)
    if method=='optimal':
        #Loop over all columns since they're likely less than rows
        for col in compDf.columns.tolist():
            #Skip if the column is missing in one of the dataframes
            if col in list(dict.fromkeys(colsNotInPre+colsNotInPost)):
                continue
            #Get the index of the nan values in that col of both dataframes
            preNan = preDf.index[pd.isnull(preDf[col])==True].tolist()
            postNan = postDf.index[pd.isnull(postDf[col])==True].tolist()
            #Get indices that are in both pre and post
            both = list(set(preNan)&set(postNan))
            #Remove items in both for notInPre and notInPost
            notInPre = list(set(preNan)-set(both))
            notInPost = list(set(postNan)-set(both))
            #Mark those cells with the relevant string
            compDf.loc[notInPre, col] = 'Not In Pre'
            compDf.loc[notInPost, col] = 'Not In Post'
            compDf.loc[both, col] = '.'
        if list(zip(*np.where(pd.isnull(compDf)))) != []:
            ct.error('Comparison dataframe contains NaNs when it shouldnt, treat'
                  'output with caution!', warning=True)
    return compDf

def missingDataCol(df):
    """Creates and fills the 'Missing Data' column of the comparison dataframe.
    """
    df['Missing Data'] = ['No']*len(df.index.tolist())
    notInPreIdxs = df[df.values == 'Not In Pre'].index.tolist()
    notInPostIdxs = df[df.values == 'Not In Post'].index.tolist()
    notInBoth = list(set(notInPostIdxs)&set(notInPreIdxs))
    noRowPreIdxs = df[df.values == 'Row Not In Pre'].index.tolist()
    noRowPostIdxs = df[df.values == 'Row Not In Post'].index.tolist()
    df.loc[notInPreIdxs+noRowPreIdxs, 'Missing Data'] = 'Missing in Pre'
    df.loc[notInPostIdxs+noRowPostIdxs, 'Missing Data'] = 'Missing in Post'
    df.loc[notInBoth, 'Missing Data'] = 'Missing in Both'
    return df

def compare():
    """Compares preDf to postDf by Post-Pre to create compDf. Adds compDf to
    filesDf.
    """
    global filesDf
    global extraCols
    for idx, preDfs in enumerate(filesDf['Pre']):
        if trackTime:
            startTime = time.time()
            noOfObs = 0
        print('Comparing', filesDf.loc[idx, 'Pre File Name'][:-18]+'...')
        dfs = []
        for (preDf, per) in preDfs:
            for (postDf, postPer) in filesDf.loc[idx, 'Post']:
                if postPer == per:
                    compDf = postDf.sub(preDf)
                    if trackTime:
                        noOfObs += len(compDf.index.tolist())*\
                                    len(compDf.columns.tolist())
                    percDf = compDf.div(preDf).mul(pd.DataFrame(100,
                                       index=compDf.index,
                                       columns=compDf.columns))
                    maxABSdiff = compDf.abs().max(axis=1)
                    maxDate = compDf.abs().idxmax(axis=1)
                    maxAsPerc = percDf.abs().max(axis=1).round(1)
                    overallABSdiff = compDf.abs().sum(axis=1)
                    #overallDiff = compDf.sum(axis=1)
                    try:
                        compDf = configureNans(preDf,postDf,compDf)
                    except MemoryError as memerr:
                        """try:
                            ct.error('Ran out of memory calculating comparison '
                                  'using the fast method. Attempting to '
                                  're-calculate using an optimised method.'
                                  '\nPlease wait...', warning=True)
                            compDf = configureNans(preDf,postDf,compDf,
                                                   method='optimal')
                        except MemoryError as memerr2:"""
                        ct.error('Ran out of memory calculating comparison '
                              'using the optimal method. Attempting to '
                              're-calculate using the low memory method.\n'
                              'This will take much longer, please wait....'
                              ,warning=True)
                        compDf = configureNans(preDf,postDf,compDf,
                                               method='lowmem')
                        """except Exception as ex2:
                            ct.error('Failed to configure NaNs.\nError Message: '+
                                  str(ex2))
                            sys.exit()"""
                    except Exception as ex:
                        ct.error('Failed to configure NaNs.\nError Message: '+
                              str(ex))
                        sys.exit()
                    compDf['Max ABS Diff (%)'] = maxAsPerc
                    compDf['Max ABS Diff'] = maxABSdiff
                    compDf['Max Diff Date'] = maxDate
                    compDf['Overall ABS Diff'] = overallABSdiff
                    #compDf['Overall Diff'] = overallDiff
                    compDf = missingDataCol(compDf)
                    extraCols = ['Overall ABS Diff',
                                 'Max ABS Diff (%)', 'Max ABS Diff',
                                 'Max Diff Date', 'Missing Data']
                    compDf.set_index(extraCols, append=True, inplace=True)
                    dfs.append((compDf, per))
        filesDf.loc[idx, 'Comp'] = dfs
        if trackTime:
            endTime = time.time()
            changeTime = round(endTime - startTime, 2)
            noOfObs = '{:,.0f}'.format(noOfObs*2)
            print('Compared %s observations in %s seconds!' % (noOfObs,
                                                               changeTime))
    
def seriesCol(df, prePost):
    """Creates the Series column for Post and Pre dataframes based on the 
    prePost variable.
    """
    if prePost not in {'BEFORE','AFTER'}:
        ct.error("seriesCol must be handed the string 'BEFORE' or 'AFTER'.")
        sys.exit()
    origIndex = df.index.names
    df['Before After'] = [prePost]*len(df.index.tolist())
    df.reset_index(inplace=True)
    df.set_index(['Before After']+origIndex, inplace=True)
    df['Series'] = list(map(':'.join,df.index.values))
    df.reset_index(inplace=True)
    df.drop(['Before After'], inplace=True, axis=1)
    df.set_index(['Series']+origIndex, inplace=True)
    return df
    
def getColWidth(df):
    """Returns the column widths for each column of a dataframe.
    """
    return [15]*len(df.index.names)+[10]*len(df.columns.tolist())

def missing_format():
    """Returns the format for missing data cells.
    """
    form = {'bold': True,
            'font_color': 'red',
            'border': 2}
    return form

def formatSheet(df, sheet, style, comparison=False):
    """Formats the xlsxwriter sheet for a given dataframe. Set comparison
    = True for the comparison sheets since they have extra columns.
    """
    ws = writer.sheets[sheet]
    wb = writer.book
    missing_index = wb.add_format({'bold': True, 'font_color': 'red',
                                   'align': 'center'})
    missing_val = wb.add_format({'bold': True, 'font_color': 'white',
                                   'border':2, 'border_color': 'black',
                                   'bg_color':'red', 'align': 'center'})
    ws.add_table(0,0,len(df.index.tolist()),
                 len(df.columns.tolist())+len(df.index.names)-1,
                 ct.createFormat(df, style, idx=True))
    if comparison:
        ws.conditional_format(1,len(df.index.names),len(df.index.tolist()),
                              len(df.columns.tolist())+len(df.index.names),
                              {'type': '3_color_scale',
                               'min_color': 'red',
                               'max_color': 'green',
                               'mid_color': 'white'})
        ws.conditional_format(1,len(df.index.names), len(df.index.tolist()),
                              len(df.columns.tolist())+len(df.index.names),
                              {'type': 'cell',
                               'criteria': 'equal to',
                               'value': '"Col Not In Post"',
                               'format': missing_index})
        ws.conditional_format(1,len(df.index.names), len(df.index.tolist()),
                              len(df.columns.tolist())+len(df.index.names),
                              {'type': 'cell',
                               'criteria': 'equal to',
                               'value': '"Col Not In Pre"',
                               'format': missing_index})
        ws.conditional_format(1,len(df.index.names), len(df.index.tolist()),
                              len(df.columns.tolist())+len(df.index.names),
                              {'type': 'cell',
                               'criteria': 'equal to',
                               'value': '"Row Not In Post"',
                               'format': missing_index})
        ws.conditional_format(1,len(df.index.names), len(df.index.tolist()),
                              len(df.columns.tolist())+len(df.index.names),
                              {'type': 'cell',
                               'criteria': 'equal to',
                               'value': '"Row Not In Pre"',
                               'format': missing_index})
        ws.conditional_format(1,len(df.index.names), len(df.index.tolist()),
                              len(df.columns.tolist())+len(df.index.names),
                              {'type': 'cell',
                               'criteria': 'equal to',
                               'value': '"Not In Pre"',
                               'format': missing_val})
        ws.conditional_format(1,len(df.index.names), len(df.index.tolist()),
                              len(df.columns.tolist())+len(df.index.names),
                              {'type': 'cell',
                               'criteria': 'equal to',
                               'value': '"Not In Post"',
                               'format': missing_val})
    for i, width in enumerate(getColWidth(df)):
        ws.set_column(i, i, width)

def outFolder():
    """Changes directories to the output folder setting one up if it doesn't 
    already exist.
    """
    os.chdir(outFol)
    fol = datetime.datetime.now().strftime("%d-%m-%Y (%H;%M;%S)")
    try:
        os.chdir(fol)
    except:
        os.mkdir(fol)
        os.chdir(fol)

def write(fileType='xlsx'):
    """Writes all outputs for files in filesDf.
    """
    global writer
    outFolder()
    if fileType not in {'csv', 'xlsx'}:
        ct.error('Unrecognised filetype chosen. Defaulting to xlsx.',
              warning=True)
        fileType = 'xlsx'
    if fileType == 'xlsx':
        for idx, file in enumerate(filesDf['Pre File Name']):
            filename = file[:-18]
            print('Creating comparison spreadsheet for', file[:-18]+'...')
            writer = pd.ExcelWriter(filename + ' Comparison.xlsx',
                                    engine='xlsxwriter')
            for per in filesDf.loc[idx, 'Periodicities'].split(','):
                if preSheet:
                    for df, p in filesDf.loc[idx, 'Pre']:
                        if p == per:
                            preDf = seriesCol(df.replace(np.nan, '.'),'BEFORE')
                if postSheet:
                    for df, p in filesDf.loc[idx, 'Post']:
                        if p == per:
                            postDf = seriesCol(df.replace(np.nan, '.'),'AFTER')
                for df, p in filesDf.loc[idx, 'Comp']:
                    if p == per:
                        compDf = df
                if preSheet:
                    preDf.to_excel(writer, sheet_name='Pre-Change ('+per+')',
                                   merge_cells=False,
                                   freeze_panes=(1,len(preDf.index.names)))
                    style = 'Table Style Medium 4'
                    formatSheet(preDf, 'Pre-Change ('+per+')', style)
                if postSheet:
                    postDf.to_excel(writer, sheet_name='Post-Change ('+per+')',
                                   merge_cells=False,
                                   freeze_panes=(1,len(postDf.index.names)))
                    style = 'Table Style Medium 7'
                    formatSheet(postDf, 'Post-Change ('+per+')', style)
                compDf.to_excel(writer, sheet_name='Difference ('+per+')',
                               merge_cells=False,
                               freeze_panes=(1,len(compDf.index.names)))
                style = 'Table Style Medium 1'
                formatSheet(compDf, 'Difference ('+per+')', style,
                            comparison=True)
            print('Saving...')
            writer.save()
    if fileType == 'csv':
        for idx, file in enumerate(filesDf['Pre File Name']):
            filename = file[:-18]
            print("Saving csv's for", file[:-18]+'...')
            for per in filesDf.loc[idx, 'Periodicities'].split(','):
                for df, p in filesDf.loc[idx, 'Comp']:
                    if p == per:
                        df.to_csv(filename + ' Difference (' + p + ').csv')
                if preSheet:
                    for df, p in filesDf.loc[idx, 'Pre']:
                        if p == per:
                            df = seriesCol(df.replace(np.nan, '.'), 'BEFORE')
                            df.to_csv(filename + ' Pre-Change (' + p + ').csv')
                if postSheet:
                    for df, p in filesDf.loc[idx, 'Post']:
                        if p == per:
                            df = seriesCol(df.replace(np.nan, '.'), 'AFTER')
                            df.to_csv(filename + ' Post-Change (' + p +').csv')
                            
                        
def configureFiles():
    """Automatically sorts the files into global lists of pre and post.
    """
    global preFiles
    global postFiles
    preFiles = []
    postFiles = []
    inputFiles = pd.DataFrame(columns=['File','Name','Year','Month','Day',
                                       'Time'])
    for file in glob.glob('*.csv'):
        r = len(inputFiles.index.tolist())
        try:
            inputFiles.loc[r,'File'] = file
            inputFiles.loc[r,'Name'] = file[:-18]
            inputFiles.loc[r,'Year'] = int(file[-13:-11])
            inputFiles.loc[r,'Month'] = int(file[-15:-13])
            inputFiles.loc[r,'Day'] = int(file[-17:-15])
            inputFiles.loc[r,'Time'] = int(file[-10:-4])
        except:
            ct.error('File: "'+file+'" is incorrectly named. Naming format must'+
                  ' match the CORD default format.\nSkipping file...')
            try:
                inputFiles.drop(r, inplace=True)
            except:
                pass
    inputFiles.sort_values(['Year','Month','Day'], ascending=[True,True,True],
                           inplace=True)
    inputFiles.reset_index(drop=True,inplace=True)
    for filename in inputFiles['Name']:
        idxs = inputFiles.index[inputFiles['Name']==filename].tolist()
        #Files in inputFiles are sorted by ascending Date and Time already
        if len(idxs) == 0: continue
        if len(idxs) == 1:
            ct.error('Only one instance of "'+filename+'" (Nothing  to compare '+
                  'to).\nFile will be skipped...', warning=True)
            inputFiles.drop(idxs,inplace=True)
            continue
        if len(idxs) == 2:
            preFiles.append(inputFiles.loc[idxs[0], 'File'])
            postFiles.append(inputFiles.loc[idxs[1], 'File'])
            inputFiles.drop(idxs, inplace=True)
            continue
        if len(idxs) > 2:
            dropFiles = ', '.join(inputFiles.loc[idxs[1:-1], 'File'].tolist())
            ct.error('Too many instances of "'+filename+'". The latest file '+
                  'will be compared to the earliest, all others will be '+
                  'ignored.\nIgnoring: '+dropFiles, warning=True)
            preFiles.append(inputFiles.loc[idxs[0], 'File'])
            postFiles.append(inputFiles.loc[idxs[-1], 'File'])
            inputFiles.drop(idxs, inplace=True)
            continue
        ct.error('Something went wrong in configureFiles.')
        sys.exit()
    
def writeGraphs():
    """Experimental feature, not currently working.
    """
    for idx, file in enumerate(filesDf['Pre File Name']):
        filename = file[:-18]
        print('Graphing results for', filename+'...')
        for per in filesDf.loc[idx, 'Periodicities'].split(','):
            for idf, p in filesDf.loc[idx, 'Comp']:
                if p == per:
                    df = idf
            for col in list(set(df.index.names)^set(extraCols)):
                for item in df.index.get_level_values(col).drop_duplicates()\
                .tolist():
                    tempDf = df[df.index.get_level_values(col).isin([item])]\
                    .reset_index()
                    tempDf.plot(kind='line', x=tempDf.columns.tolist(),
                                y='Overall ABS Diff')
                    plt.show()
        print('Saving...')

def runTask(lZeros=True, createGraphs=False):
    """Main task for running the script.
    """
    global filesDf, inpFol, outFol
    if trackTime: start = time.time()
    #Ignore the peformance warning that will sometimes appear in configureNans
    #Remove for debugging.
    warnings.simplefilter(action='ignore',
                          category=pd.errors.PerformanceWarning)
    filesDf = pd.DataFrame(columns=['Pre File Name', 'Post File Name',
                                    'Periodicities', 'Pre', 'Post', 'Comp'])
    (inpFol, outFol) = ct.setupFilepaths()
    os.chdir(inpFol)
    configureFiles()
    readPreChange(lZeros)
    readPostChange(lZeros)
    compare()
    write(fileType)
    if createGraphs:
        writeGraphs()
    if trackTime:
        end = time.time()
        change = round(end - start,2)
        print('Took %s seconds to create, format, and save all comparisons.'
              % change)
        
#================================USER OPTIONS=================================
#Convert leading zeros if the file has been previously modified in excel?
leadingZeros = True
#Choose the file extension of the output. Options are 'xlsx' or 'csv'
fileType = 'xlsx'
#Create Pre sheets in the output? False reduces file size.
preSheet = False
#Create Post sheets in the output? False reduces file size.
postSheet = False
#Track the time taken for each comparison?
trackTime = True
#=============================================================================

createGraphs = False
runTask(leadingZeros, createGraphs)
os.chdir(inpFol)
ct.done()