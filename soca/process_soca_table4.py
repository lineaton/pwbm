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
    elif year < 2000:
        if year == 1999:
            start_row_indeces = [11, 49, 87, 125, 163]
        elif year in [1997, 1998]:
            start_row_indeces = [13, 52, 91, 130, 169]
        elif year == 1997:
            start_row_indeces = [14, 55, 96, 136, 177]
    
    results = pd.DataFrame()
    #process from each of the indeces and put into our categories of: 
    #Data, TableID, Year, HoldingPeriod, TransactionType, AssetType, AGI, Series, Value
    # row_range = range(1, 9)

    #18 rows of data from each table, 5 tables
    table_height = 90
    # output = [["" for x in row_range] for y in range(table_height)]
    # for (row, col) in product(range(table_height), row_range):
    #     if is_number(str(sheet.row(row)[col].value)):
    #         output[row][col] = float(to_number(str(sheet.row(row)[col].value)))
    #     else:
    #         output[row][col] = clean_column_name(str(sheet.row(row)[col].value))
    
    # output = np.asarray(output)

    # based on row, changes after each start_row_index
    table_ids = ['A', 'B', 'C', 'D', 'E']
    asset_type = ["AllAsset", "CorporateStock", "BondsAndOtherSecurities", "RealEstate", "Other"]

    # based on row, in each pair of subtables, first in pair is short term, second is long term
    holding_periods = ["ShortTerm", "LongTerm"]

    # based on row in subtable 
    agi_categories = ["AllReturns", "AdjustedGrossDeficit", "Under 20,000", \
    "20,000 under 50,000", "50,000 under 100,000", "100,000 under 200,000", \
    "200,000 under 500,000", "500,000 under 1,000,000", "1,000,000 or more"]

    # based on column: all is [1], gain: [2-4], loss [5-7] 
    transaction_types = ["AllTransactions", "GainTransaction", "GainTransaction", "GainTransaction", "LossTransaction", "LossTransaction", "LossTransaction"]

    # based on column: trans_inds: returns, trans_inds + 1: trans, 4: gain, 7: loss
    series = ["NumberOfReturns", "NumberOfReturns", "NumberOfTransactions", "Gain", "NumberOfReturns", "NumberOfTransactions", "Loss"]

    # generate subtable dataframe
    result = []

    # iterate through each table pair
    for table_no in range(len(start_row_indeces)):
        shortTermIndex = start_row_indeces[table_no]
        longTermIndex = start_row_indeces[table_no] + space
        
        table_inds = [start_row_indeces[table_no], start_row_indeces[table_no] + space]

        #iterate through rows in each table for short term
        for table in range(len(table_inds)):
            for row in range(len(agi_categories)):
                rowInd = table_inds[table] + row - 1
                holdingPeriod = holding_periods[table]
                for col in range(len(transaction_types)):
                    colInd = col + 1
                    value = sheet.cell_value(rowInd, colInd)
                    result.append({
                        "Value": value,
                        "Series": series[col],
                        "AGI": agi_categories[row],
                        "AssetType": asset_type[table_no],
                        "TransactionType": transaction_types[col],
                        "HoldingPeriod": holdingPeriod,
                        "Year": year,
                        "TableID": "Table2" + table_ids[table_no],
                        "Data":'SOIIndividual'
                    })

    subtables = pd.DataFrame(result).loc[:, ["Data", "TableID", "Year", "HoldingPeriod", "TransactionType", \
        "AssetType", "AGI", "Series", "Value"]]
    results = pd.concat([results, subtables])
    return results


def process_ip_table(directory):
    wb = xlrd.open_workbook(filename = directory, formatting_info=True)
    sheet = wb.sheet_by_index(0)
    # determine where each table starts
    start_row = 8 - 1
    start_col = 1
    
    results = pd.DataFrame()
    #process from each of the indeces and put into our categories of: 
    #Data, TableID, Year, HoldingPeriod, TransactionType, AssetType, AGI, Series, Valu

    # based on row, changes after each start_row_index
    table_ids = ["N/A"]
    asset_type = ["N/A"]

    # based on row, in each pair of subtables, first in pair is short term, second is long term
    holding_periods = ["N/A"]

    # based on row in subtable 
    agi_categories = ["AllReturns", "AdjustedGrossDeficit and Under 5,000", "5,000 under 10,000", "10,000 under 15,000", "15,000 under 20,000", \
    "20,000 under 25,000", "25,000 under 30,000", "30,000 under 40,000", "40,000 under 50,000", "50,000 under 75,000", "75,000 under 100,000", "100,000 under 200,000", \
    "200,000 under 500,000", "500,000 under 1,000,000", "1,000,000 under 1,500,000", "1,500,000 under 2,000,000", "2,000,000 under 5,000,000", "5,000,000 under 10,000,000", "10,000,000 or more"]

    transaction_types = ["N/A"]

    # based on column: trans_inds: returns, trans_inds + 1: trans, 4: gain, 7: loss
    series = ["NumberOfReturns", "NumberOfReturns", "NumberOfTransactions", "Gain", "NumberOfReturns", "NumberOfTransactions", "Loss"]

    # generate subtable dataframe
    result = []

    # iterate through each table pair
    # for table_no in range(len(start_row_indeces)):
        

    #     #iterate through rows in each table for short term
    #     for table in range(len(table_inds)):
    #         for row in range(len(agi_categories)):
    #             rowInd = table_inds[table] + row - 1
    #             holdingPeriod = holding_periods[table]
    #             for col in range(len(transaction_types)):
    #                 colInd = col + 1
    #                 value = sheet.cell_value(rowInd, colInd)
    #                 result.append({
    #                     "Value": value,
    #                     "Series": series[col],
    #                     "AGI": agi_categories[row],
    #                     "AssetType": asset_type[table_no],
    #                     "TransactionType": transaction_types[col],
    #                     "HoldingPeriod": holdingPeriod,
    #                     "Year": year,
    #                     "TableID": "Table2" + table_ids[table_no],
    #                     "Data":'SOIIndividualPanel'
    #                 })

    # subtables = pd.DataFrame(result).loc[:, ["Data", "TableID", "Year", "HoldingPeriod", "TransactionType", \
    #     "AssetType", "AGI", "Series", "Value"]]
    # results = pd.concat([results, subtables])
    # return results








def process_soca_table_2():
    # load interfaces
    # _interface_paths = json.load(open(os.path.join('..','..','.interface_paths.json')))
    # # if cache reside with the repository instead of an outside drive.... this is necessary
    # os.chdir(os.path.join('..','..'))
    
    # if not os.path.exists(os.path.join(_interface_paths["cache"], 'Interfaces')):
    #     os.makedirs(os.path.join(_interface_paths["cache"], 'Interfaces'))
        
    # directory = os.path.join(_interface_paths['soca'], 'soca_table_2')
    directory = os.path.join("Desktop", "pwbm", "soca", 'soca_table_2')
    print(directory)
    process_xlsx_conversion(directory) 
    
    results = pd.DataFrame()

    file_list_1 = ['07in02ab.xls','08in02soca.xls',\
                 '09in02soca.xls','1002insoca.xls','1102insoca.xls','1202insoca.xls',\
                #  '97soca2a.xls',
                '98in6ab.xls','98in2ab.xls','99in02ab.xls']
    
    # individual panel
    file_list_2 = ['07in02ai.xls', '99-03in02ai.xls']

    
    for filename in file_list_1:
        # for mac error
        if filename != '.DS_Store':
            print ("------")
            print(filename)
            # file_path = (os.path.join(_interface_paths['soca'], 'soca_table_2',filename))
            file_path = (os.path.join(directory, filename))
            year = filename_year(filename)
            if filename =='98in6ab.xls':
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

    results.to_csv(os.path.join(directory, 'table_2.csv'), index=False)

if __name__ == "__main__":
    process_soca_table_2()