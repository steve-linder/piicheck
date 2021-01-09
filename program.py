import os
import sys
import pandas as pd
import re
import PySimpleGUI as sg
import logging
from datetime import datetime

# TODO convert rest to pysimple gui
# TODO old excel files?
# TODO pysimplegui file picker
# TODO Search for filenames and column headers
# TODO Whitelist folders and files


# cd /c/python38/source/steve/test
# usage: python program.py 'Q:\Hospitality'


def main():
    walk_dir=get_path()
    numfiles = 0
    numxfiles = 0
    checkpii=False  #Looks for SSN's 
    checkname=False #Looks for names with password in it
    checknewest=True #Identifies the Newest File
    newest_modified_date=0
    max_size=10000000
    outcome=False
    logging.basicConfig(filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')
    
    for root, subdirs, files in os.walk(walk_dir):
        for filename in files:
            file_path = os.path.join(root, filename)
            filename, file_extension = os.path.splitext(file_path)
            print(file_path)
            numfiles += 1
            if checknewest==True:
                last_modified_date=os.path.getmtime(file_path)
                if newest_modified_date < last_modified_date:
                    newest_modified_date = last_modified_date
                print(filename,datetime.fromtimestamp(last_modified_date).strftime('%Y-%m-%d %H:%M:%S'))
            if checkname==True:
                if "password" in filename.lower():
                    logging.warning('SUSPICIOUS FILE :'+file_path)
            if file_extension == '.xlsx' and checkpii==True:
                size = os.stat(file_path).st_size
                if (size>max_size):
#                    print('FILE TOO LARGE', filename, size)
                    logging.warning('FILE TOO LARGE :'+file_path)
                else:
                    numxfiles += 1
                try:
                    df = pd.read_excel(file_path)
                except:
                    logging.warning('ERROR OPENING FILE '+file_path)
                else:   
                    pii_outcome=analyze_excel(df, file_path)
                    if (pii_outcome==True):
                        outcome=True 

    print('Number of files: ', numfiles)
    print('Number of Excel files: ', numxfiles)
    if (checknewest==True):            
        print(datetime.fromtimestamp(newest_modified_date).strftime('%Y-%m-%d %H:%M:%S'))
    if (outcome==True):
        print('WARNING!!!-PII Data found!!! Check logs now!!!')

def get_path():
    text = sg.popup_get_folder('Please enter a folder name')
    return text

def analyze_excel(df, file_path):
    print('Reading: ',file_path)
    columns = len(df.columns)
    rows = len(df.index)

    if rows > 10:
        rows = 10
    
    pat = re.compile('^\d{3}-\d{2}-\d{4}$')

    ssnfound = False
    for x in range(0, rows):
        for y in range(0, columns):
            celldata = str(df.iloc[x, y])
            resp = re.match(pat, celldata)
            if (bool(resp) == True):
                ssnfound = True
                break
    if (ssnfound == True):
        logging.warning('Pii data found in :'+file_path)
    
        
    return ssnfound
         

if __name__ == '__main__':
    main()
