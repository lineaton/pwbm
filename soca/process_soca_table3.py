# -*- coding: utf-8 -*-
'''
    File name: process_soca_table3.py
    Author: Eaton Lin
    Date created: 6/12/2020
    Python version: 2.7
    Description: Script to parse through and process from SOCA data as classified by:
        Selected Asset Type and Month of Sale
'''

import pandas as pd
import xlrd
import os
from itertools import product
import numpy as np

from utilities import   \
    is_number,          \
    to_number,          \
    to_camelcase,       \
    get_subtables_2,    \
    get_subtables_3,    \
    clean_column_name,  \
    get_subtables,      \
    filename_year,      \
    process_xlsx_conversion

def process_soca_table(directory, year):
    wb = xlrd.open_workbook(filename = directory, formatting_info=True)
    sheet = wb.sheet_by_index(0)
    space = 15
    # determine where each table starts
    start_row_indeces = []
    if year == 2007:
        start_row_indeces = [9, 50, 91, 132, 173]
    elif year in [2008, 2009]:
        start_row_indeces = [9, 52, 95, 139, 183]
    elif year in [2010, 2011, 2012]:
        start_row_indeces = [9, 47, 85, 123, 161]
    elif year == 1999:
            start_row_indeces = [11, 51, 91, 131, 171]
    elif year in [1997, 1998]:
        start_row_indeces = [10, 49, 88, 127, 166]
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
    months = ["AllReturns", "January", "February", \
    "March", "April", "May", \
    "June", "July", "August", \
    "September", "October", "November", \
    "December", "Not determinate"]

    # based on column: all is [1], gain: [2-4], loss [5-7] 
    transaction_types = ["GainTransaction", "GainTransaction", "GainTransaction", "GainTransaction","LossTransaction", "LossTransaction", "LossTransaction", "LossTransaction"]

    # based on column: trans_inds: returns, trans_inds + 1: trans, 4: gain, 7: loss
    series = ["NumberOfTransactions", "SalesPrice", "Basis", "Gain", "NumberOfTransactions", "SalesPrice", "Basis", "Loss"]

    # generate subtable dataframe
    result = []

    # iterate through each table pair
    for table_no in range(len(start_row_indeces)):
        table_inds = [start_row_indeces[table_no], start_row_indeces[table_no] + space]
        for table in range(len(table_inds)):
            for row in range(len(months)):
                rowInd = table_inds[table] + row - 1
                holdingPeriod = holding_periods[table]
                for col in range(len(transaction_types)):
                    colInd = col + 1
                    value = sheet.cell_value(rowInd, colInd)
                    result.append({
                        "Value": value,
                        "Series": series[col],
                        "Month": months[row],
                        "AssetType": asset_type[table_no],
                        "TransactionType": transaction_types[col],
                        "HoldingPeriod": holdingPeriod,
                        "Year": year,
                        "TableID": "Table3" + table_ids[table_no],
                        "Data":'SOIIndividual'
                    })
        

    subtables = pd.DataFrame(result).loc[:, ["Data", "TableID", "Year", "HoldingPeriod", "TransactionType", \
        "AssetType", "Month", "Series", "Value"]]
    results = pd.concat([results, subtables])
    return results


def process_soca_table_3():
    directory = os.path.join('soca_table_3')
    print(directory)
    process_xlsx_conversion(directory) 
    
    results = pd.DataFrame()

    file_list = ['07in03ab.xls','08in03soca.xls',\
                 '09in03soca.xls','1003insoca.xls','1103insoca.xls','1203insoca.xls',\
                #  '97soca3a.xls',
                '98in7ab.xls','98in3ab.xls','99in03ab.xls']

    
    for filename in file_list:
        # for mac error
        if filename != '.DS_Store':
            print ("------")
            print(filename)
            file_path = (os.path.join(directory, filename))
            year = filename_year(filename)
            if filename =='98in7ab.xls':
                year = 1997
            result_df = process_soca_table(file_path, year)
            results = pd.concat([results, result_df])

    results.to_csv(os.path.join(directory, 'table_3.csv'), index=False)

if __name__ == "__main__":
    process_soca_table_3()