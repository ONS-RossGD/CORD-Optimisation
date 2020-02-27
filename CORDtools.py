# -*- coding: utf-8 -*-
"""
Common functions used across CORD Optimisation scripts.
"""

import sys, os, glob, datetime
import tkinter as tk, pandas as pd
from zipfile import ZipFile
from tkinter.filedialog import askdirectory
    
def error(msg, warning=False):
    """Prints an error message in red.
    """
    if warning:
        sys.stdout.write("\033[1;33m")
        print('WARNING:', msg)
    else:
        sys.stdout.write("\033[1;31m")
        print('ERROR:', msg)
    sys.stdout.write("\033[0;0m")

def setupFilepaths():
    root = tk.Tk()
    inpFol = askdirectory(title="Choose an INPUT folder")
    outFol = askdirectory(title="Choose an OUTPUT folder")
    root.destroy()
    if inpFol == '':
        error('An INPUT folder must be selected.')
        sys.exit()
    if outFol == '':
        error('An OUTPUT folder must be selected.')
        sys.exit()
    return (inpFol, outFol)

def createFormat(df, style, idx=False):
    """Creates the xlsxwriter format for a given dataframe and returns it as a
    string.
    """
    cols = []
    if idx:
        for col in df.index.names:
            cols.append({'header':col})
    for col in df.columns.tolist():
        cols.append({'header':col})
    form = {'columns':cols, 'style': style}
    return form

def done():
    sys.stdout.write("\033[1;32m")
    print('Done!')
    sys.stdout.write("\033[0;0m")
    
def unzipFiles(file):
    """Unzips the specified file.
    """
    fileTitle = os.path.splitext(file)[0]
    print('Unzipping', fileTitle, '...')
    with ZipFile(file, 'r') as zipObj:
        zipObj.extractall(fileTitle)
        
def splitCordString(s):
    """Split up a CORD Output string (in format {*** = ***}) into individual
    strings contained in list splitString. Returns splitString.
    """
    s = s.replace('{classification', ' classification')
    s = s.replace('} ','')
    s = s.replace('}','')
    s = s.replace("'", '')
    s = s.replace(' {', '{')
    splitString = s.split('{')
    while '' in splitString:
        splitString.remove('')
    for i, ss in enumerate(splitString):
        splitString[i] = ss.strip()
    return splitString

def createOutFolder(outFol):
    """Changes directories to the output folder setting one up if it doesn't 
    already exist. Returns the created folder path.
    """
    returnPath = os.getcwd()
    os.chdir(outFol)
    fol = datetime.datetime.now().strftime("%d-%m-%Y (%H;%M;%S)")
    try:
        os.chdir(fol)
    except:
        os.mkdir(fol)
        os.chdir(fol)
    outFolder = os.getcwd()
    os.chdir(returnPath)
    return outFolder
    
def unzipConfigReports(fol):
    """Unzips all Configuration Reports in the specified folder. Adds the
    meta data to a pandas DataFrame configReports which is returned.
    """
    os.chdir(fol)
    configReports = pd.DataFrame(columns=['Config Report',
                                          'Statistical Activity',
                                          'Mode', 'Date'])
    for file in glob.glob('*.zip'):
        if '_Config_Report_' not in file:
            error('"'+file+'" is not a Config Report and will be ignored.',
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
        with ZipFile(file, 'r') as zipObj:
             zipObj.extractall(fileTitle)
    return configReports