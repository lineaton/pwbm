# -*- coding: utf-8 -*-
'''
    File name: process_soca_table2.py
    Author: Eaton Lin
    Date created: 6/3/2020
    Python version: 2.7
    Description: Script to parse through and process from SOCA data as classified by:
        Size of Adjusted Gross Income and Selected Asset Type
'''

import pandas as pd
import json
import xlrd
import os
from itertools import product
import numpy as np
import win32com.client

from utilities import   \
    is_number,          \
    to_number,          \
    to_camelcase,       \
    get_subtables_2,    \
    get_subtables_3,    \
    clean_column_name,  \
    get_subtables,      \
    filename_year

def process_xlsx_conversion(directory):
    conversion_list = []
    for path, subdirs, files in os.walk(directory):
        for name in files:
            if name.endswith('xlsx') :
                conversion_list.append(os.path.join(path, name))
    
    xl = win32com.client.Dispatch("Excel.Application")
    xl.DisplayAlerts = False
    for item in conversion_list: 
        if not os.path.exists(os.path.join(item[:-1])) :
            wb = xl.Workbooks.Open(os.path.join(os.getcwd(), item))
            wb.SaveAs(os.path.join(os.getcwd(), item.replace('.xlsx', '.xls')), FileFormat = 56)
            wb.Close()
    
    xl.Quit()

def process_soca_table(directory, year):
    wb = xlrd.open_workbook(filename = directory, formatting_info=True)
    sheet = wb.sheet_by_index(0)
    space = 16
    # determine where each table starts
    start_row_indeces = []
    if year in [2007, 2010, 2011]:
        start_row_indeces = [9, 45, 81, 117, 153]
    elif year in [2008, 2009]:
        start_row_indeces = [9, 52, 95, 139]
        if year == 2009:
            start_row_indeces.append(184)
        else:
            start_row_indeces.append(183)
    elif year == 2012:
        start_row_indeces = [9, 53, 97, 142, 187]
    elif year == 1999:
            start_row_indeces = [11, 49, 87, 125, 163]
    elif year in [1997, 1998]:
        start_row_indeces = [10, 47, 84, 121, 158]
    else:
        raise ValueError("Year out of bounds.")    
    
    results = pd.DataFrame()
    #process from each of the indeces and put into our categories of: 
    #Data, TableID, Year, HoldingPeriod, TransactionType, AssetType, AGI, Series, Value
    # row_range = range(1, 9)

    # based on row, changes after each start_row_index
    table_ids = ['A', 'B', 'C', 'D', 'E']
    asset_type = ["AllAsset", "CorporateStock", "BondsAndOtherSecurities", "RealEstate", "Other"]

    # based on row, in each pair of subtables, first in pair is short term, second is long term
    holding_periods = ["ShortTerm", "LongTerm"]

    # based on row in subtable 
    st_time_held_cats = ["AllReturns", "Under 1 month", "1 month under 2 months", \
    "2 months under 3 months", "3 months under 4 months", "4 months under 5 months", \
    "5 months under 6 months", "6 months under 7 months", "7 months under 8 months", \
    "8 months under 9 months", "9 months under 10 months", "10 months under 11 months", \
    "11 months under 12 monthss", "1 year or more", "Holding period not determinable"]

    lt_time_held_cats = ["AllReturns", "Under 18 months", "18 months under 2 years", \
    "2 years under 3 years", "3 years under 4 years", "4 years under 5 years", \
    "5 years under 10 years", "10 years under 15 years", "15 years under 20 years", \
    "20 years or more", "Holding period not determinable"]

    # based on column: all is [1], gain: [2-4], loss [5-7] 
    transaction_types = ["GainTransaction", "GainTransaction", "GainTransaction", "GainTransaction","LossTransaction", "LossTransaction", "LossTransaction", "LossTransaction"]

    # based on column: trans_inds: returns, trans_inds + 1: trans, 4: gain, 7: loss
    series = ["NumberOfTransactions", "SalesPrice", "Basis", "Gain", "NumberOfTransactions", "SalesPrice", "Basis", "Loss"]

    # generate subtable dataframe
    result = []

    # iterate through each table pair
    for table_no in range(len(start_row_indeces)):
        shortTermIndex = start_row_indeces[table_no]
        longTermIndex = start_row_indeces[table_no] + space

        for row in range(len(st_time_held_cats)):
            rowInd = shortTermIndex + row - 1
            holdingPeriod = holding_periods[0]
            for col in range(len(transaction_types)):
                colInd = col + 1
                value = sheet.cell_value(rowInd, colInd)
                result.append({
                    "Value": value,
                    "Series": series[col],
                    "TimeHeld": st_time_held_cats[row],
                    "AssetType": asset_type[table_no],
                    "TransactionType": transaction_types[col],
                    "HoldingPeriod": holdingPeriod,
                    "Year": year,
                    "TableID": "Table4" + table_ids[table_no],
                    "Data":'SOIIndividual'
                })

        for row in range(len(lt_time_held_cats)):
            rowInd = longTermIndex + row - 1
            holdingPeriod = holding_periods[1]
            for col in range(len(transaction_types)):
                colInd = col + 1
                value = sheet.cell_value(rowInd, colInd)
                result.append({
                    "Value": value,
                    "Series": series[col],
                    "TimeHeld": lt_time_held_cats[row],
                    "AssetType": asset_type[table_no],
                    "TransactionType": transaction_types[col],
                    "HoldingPeriod": holdingPeriod,
                    "Year": year,
                    "TableID": "Table4" + table_ids[table_no],
                    "Data":'SOIIndividual'
                })

    subtables = pd.DataFrame(result).loc[:, ["Data", "TableID", "Year", "HoldingPeriod", "TransactionType", \
        "AssetType", "TimeHeld", "Series", "Value"]]
    results = pd.concat([results, subtables])
    return results

def process_soca_table_4():
    # load interfaces
    # _interface_paths = json.load(open(os.path.join('..','..','.interface_paths.json')))
    # # if cache reside with the repository instead of an outside drive.... this is necessary
    # os.chdir(os.path.join('..','..'))
    
    # if not os.path.exists(os.path.join(_interface_paths["cache"], 'Interfaces')):
    #     os.makedirs(os.path.join(_interface_paths["cache"], 'Interfaces'))
        
    # directory = os.path.join(_interface_paths['soca'], 'soca_table_2')
    directory = os.path.join('soca_table_4')
    print(directory)
    process_xlsx_conversion(directory) 
    
    results = pd.DataFrame()

    file_list_1 = ['07in04ab.xls','08in04soca.xls',\
                 '09in04soca.xls','1004insoca.xls','1104insoca.xls','1204insoca.xls',\
                #  '97soca2a.xls',
                '98in8ab.xls','98in4ab.xls','99in04ab.xls']
    
    # individual panel
    file_list_2 = ['07in03tt.xls', '99-03in03tt.xls']

    
    for filename in file_list_1:
        # for mac error
        if filename != '.DS_Store':
            print ("------")
            print(filename)
            # file_path = (os.path.join(_interface_paths['soca'], 'soca_table_2',filename))
            file_path = (os.path.join(directory, filename))
            year = filename_year(filename)
            if filename =='98in8ab.xls':
                year = 1997
            result_df = process_soca_table(file_path, year)
            results = pd.concat([results, result_df])
            
    # for filename in file_list_2:
    #     if filename != '.DS_Store':
    #         print ("------")
    #         print(filename)
    #         file_path = (os.path.join(_interface_paths['soca'], 'soca_table_2',filename))
    #         if filename == '04-07in05st.xls':
    #             file_path = os.path.join('SOI','SOCA', filename)
    #         result_df = process_widetable(file_path)
    #         results = pd.concat([results, result_df])
    
    # results.to_csv(os.path.join(_interface_paths['cache'],'Interfaces', 'table_2.csv'), index=False)

    results.to_csv(os.path.join(directory, 'table_4.csv'), index=False)

if __name__ == "__main__":
    process_soca_table_4()