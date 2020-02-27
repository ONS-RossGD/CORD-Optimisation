# -*- coding: utf-8 -*-
"""
Common functions used across CORD Optimisation scripts.
"""

import sys, os
from zipfile import ZipFile
import tkinter as tk
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