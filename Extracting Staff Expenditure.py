# -*- coding: utf-8 -*-
"""
Created on Mon Sep 11 13:46:15 2017

@author: Iwan.Williams
"""

import itertools
import glob
import shutil
from openpyxl import load_workbook
from pandas import DataFrame
import pandas as pd
import os
import numpy as np

#excel spreadsheets collected from individual folders in our internal drive 
#into a single folder.
for folderName, subfolders, filenames in os.walk(
    r'\\rdtaxserver\company\R&D\Reports'):

    for summary in filenames:
        if 'Costs Summary' or 'Cost Summary' in summary:
                    try: shutil.copy(os.path.join(folderName, summary), 
                    r'C:\Users\iwan.williams\Documents\Working Python Files\
                    Cost summaries')
                    except: 
                        print(folderName, summary)


def get_data(sheet):
    for row in sheet.values:
        row_it = iter(row)
            for cell in row_it:
                if cell is not None:
                    yield itertools.chain((cell,), row_it)
                    break


def squeeze_nan(x):
    original_columns = x.index.tolist()

    squeezed = x.dropna()
    squeezed.index = [original_columns[n] for n in range(squeezed.count())]

    return squeezed.reindex(original_columns, fill_value=np.nan)


path =r'C:\Users\iwan.williams\Documents\Working Python Files\Cost summaries'
#specify the path we're extracting information from
all_files = glob.glob(path + "/*.xls?")
# specifying that we want to obtain each file in the destination which is an 
# excel file. 
# All excel files are needed therefore the extension can be xlsx or xlsm because
# the ? regex wildcard is used.
# list of identifiers of the expenditure from the start to the end of the data.
staff_labels = ['''"Staff Costs":"Total  Staff Costs Relating to R&D"''',
'''"Staff Costs":"Total Staff Costs Relating to R&D"''',
'''"Staff Costs":"Total of Staff Costs Relating to R&D"''',
'''"Staff":"Total staff costs relating to R&D"''']

frame = pd.DataFrame()
extracted_costs = pd.DataFrame()
df = pd.DataFrame()
extracted_cost_list = []
#below we're looping over each file in the destination.


for file in all_files:
#the load_workbook function is used to activate each workbook in turn for 
#us to manipulate them in memory. We only want the raw data and we don't want 
#any excel functions.
    try: workbook = load_workbook(file, data_only=True)
    except: pass
#not all workbooks are successfully iterated over, some worksheets throw errors. 
#We iterate over each worksheet in the workbook, which are unspecified, 
#therefore a nested for loop is ideal to ensure all worksheets are identified.
    for sheet in workbook.worksheets:
        #the get_data function is used to loop over the cells in the worksheet
        #to obtain the data.
        try: 
            extracted_sheets=DataFrame(get_data(sheet))  #empty columns are 
            #dropped using the dropna function
            extracted_sheets.dropna(axis=0, thresh=2, inplace = True)
            #we also use the squeeze nan function to get rid of any spaces in 
            #the first colummns. 
            #the data has been irregularly entered, but is always in the same 
            #format, therefore by squeezing the data to the first columns we 
            #will have the data in the correct columns.
            df = extracted_sheets.apply(squeeze_nan, axis=1)
            #the index is set on the first column and is stripped of any whitespace
            df.set_index(df.columns[0], inplace = True)
            df.index.str.strip()
            #we call on the Staff_Costs list to retrieve the labels to extract 
            #the staff costs
            extracted_costs = df.loc[staff_labels[0]].copy()
            #the source filename and worksheet is added to be clear where the 
            #data is sourced from
            extracted_costs['filename'] = os.path.basename(file)
            extracted_costs['Worksheet'] = sheet
            #the data pattern consists of name, job title, cost, apportionment, 
            #and calculation of cost * apportionnment %. only the first four 
            #columns are necessary
            extracted_costs = extracted_costs[[1, 2, 3, 'filename', 'Worksheet']]
            extracted_cost_list.append(extracted_costs)
        except: print(folderName, summary)

concatenated_costs = pd.concat(extracted_cost_list)
#Reindexed to ensure index is set correctly and double checking to ensure 
#whitespace isn't present.
concatenated_costs.set_index(concatenated_costs.columns[0], inplace = True)
concatenated_costs.index.str.strip()
concatenated_costs.columns = concatenated_costs.columns.str.title().str.strip()
#labels used to locate the data are removed
concatenated_costs.drop("Staff Costs", inplace = True)
concatenated_costs.drop("Staff", inplace = True)
concatenated_costs.drop("Total  Staff Costs Relating to R&D", inplace = True)
concatenated_costs.drop("Total Staff Costs Relating to R&D", inplace = True)
concatenated_costs.drop("Total of Staff Costs Relating to R&D", inplace = True)
concatenated_costs.drop("Total staff costs relating to R&D", inplace = True)


#regex to identify client ID within file destinations. This looks for three or 
#four digits, avoids the 2013-2017 periods.
cif_regex= r'(?<!\d)((?!201[3-7])\d{4}|\d{3})(?!\d)'
frame['Cifs']=frame['Filename'].str.findall(cif_regex)
#the IDs found then need to be exploded into seperate rows with the 
#corresponding data. 
#I have found a solution at 
#https://stackoverflow.com/questions/12680754/split-explode-pandas-dataframe-string-entry-to-separate-rows 
#BUT haven't figured out how to use it correctly
os.chdir(r'C:\Users\iwan.williams\Documents')
csv = concatenated_costs.to_csv("Staff Costs.csv", sep=',', encoding='utf-8')
