# -*- coding: utf-8 -*-
"""
Created on Thu Jan  4 17:31:05 2018

@author: Iwan
"""

import itertools
import glob
from openpyxl import load_workbook
from pandas import DataFrame
import pandas as pd
import os
import numpy as np
def get_data(ws):
        for row in ws.values:
            row_it=iter(row)
            for cell in row_it:
                if cell is not None:
                    yield itertools.chain((cell,), row_it)
                    break
    
def squeeze_nan(x):
    original_columns=x.index.tolist()

    squeezed=x.dropna()
    squeezed.index=[original_columns[n] for n in range(squeezed.count())]

    return squeezed.reindex(original_columns, fill_value=np.nan)

cif_regex= r'(?<!\d)((?!201[3-7])\d{4}|\d{3})(?!\d)(?!1415)(?!1516)(?!1617)'
path =r'I:\Working Python Files\Cost summaries'
# specify the path we're extracting information from
all_files=glob.glob(path + "/*.xlsx")
# specifying that we want to obtain each file in the destination which is an 
# excel file. 
# All excel files are needed therefore the extension can be xlsx or xlsm because
# the ? regex wildcard is used.
# list of identifiers of the expenditure from the start to the end of the data.
labels=[["Staff Costs":"Total  Staff Costs Relating to R&D"],
["Staff Costs":"Total Staff Costs Relating to R&D"],
["Staff Costs":"Total of Staff Costs Relating to R&D"],
["Staff":"Total staff costs relating to R&D"], ["Total staff costs relating to R&D"]]

frame=pd.DataFrame()
frame=pd.DataFrame()
ws_list=[]

Staff4_=pd.DataFrame()
removed_nan=pd.DataFrame()
text_extract=pd.DataFrame()

for summary in all_files:
    try: wb=load_workbook(summary, data_only=True)
    except: pass
    for sheet in wb.worksheets:
        ws=sheet
        raw_d=DataFrame(get_data(ws))  
        try: 
            raw_d.dropna(axis=0, thresh=2, inplace=True)
            removed_nan=raw_d.apply(squeeze_nan, axis=1)
            removed_nan.set_index(removed_nan.columns[0], inplace=True)
            removed_nan.index.str.strip()
            text_extract=removed_nan.loc[labels].copy()
            text_extract['filename']=os.path.basename(summary)
            text_extract['Worksheet']=ws
            text_extract=text_extract[[1, 2, 3, 4, 'filename', 'Worksheet']]
            ws_list.append(text_extract)
        except: pass

concatenated_costs=pd.concat(ws_list)
# Reindexed to ensure index is set correctly and double checking to ensure 
# whitespace isn't present.
concatenated_costs.set_index(concatenated_costs.columns[0], inplace=True)
concatenated_costs.index.str.strip()
concatenated_costs.columns=concatenated_costs.columns.str.title().str.strip()
# labels used to locate the data are removed
concatenated_costs.drop("Staff Costs", inplace=True)
concatenated_costs.drop("Staff", inplace=True)
concatenated_costs.drop("Total  Staff Costs Relating to R&D", inplace=True)
concatenated_costs.drop("Total Staff Costs Relating to R&D", inplace=True)
concatenated_costs.drop("Total of Staff Costs Relating to R&D", inplace=True)
concatenated_costs.drop("Total staff costs relating to R&D", inplace=True)


# regex to identify client ID within file destinations. This looks for three or 
# four digits, avoids the 2013-2017 periods.
cif_regex= r'(?<!\d)((?!201[3-7])\d{4}|\d{3})(?!\d)'
concatenated_costs['Cifs']=concatenated_costs['Filename'].str.findall(cif_regex)
# the IDs found then need to be exploded into seperate rows with the 
# corresponding data. 

def tidy_split(df, column, sep=',', keep=False):
    """
    Split the values of a column and expand so the new DataFrame has one split
    value per row. Filters rows where the column is missing.

    Params
    ------
    df : pandas.DataFrame
        dataframe with the column to split and expand
    column : str
        the column to split and expand
    sep : str
        the string used to split the column's values
    keep : bool
        whether to retain the presplit value as it's own row

    Returns
    -------
    pandas.DataFrame
        Returns a dataframe with the same columns as `df`.
    """
    indexes = list()
    new_values = list()
    df = df.dropna(subset=[column])
    for i, presplit in enumerate(df[column].astype(str)):
        values = presplit.split(sep)
        if keep and len(values) > 1:
            indexes.append(i)
            new_values.append(presplit)
        for value in values:
            indexes.append(i)
            new_values.append(value)
    new_df = df.iloc[indexes, :].copy()
    new_df[column] = new_values
    return new_df
          
concatenated_costs=tidy_split(concatenated_costs, 'Cifs', sep=',')
