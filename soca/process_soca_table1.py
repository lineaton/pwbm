# -*- coding: utf-8 -*-
"""
Created on Tue May  7 16:48:21 2019

@author: Xiaoyue Sun
"""

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

    
def process_xlsx_conversion(dir_path):
    conversion_list = []
    for path, subdirs, files in os.walk(dir_path):
        for name in files:
            if name.endswith('.xlsx'):
                conversion_list.append(os.path.join(path, name))

    xl = win32com.client.Dispatch("Excel.Application")
    xl.DisplayAlerts = False
    for item in conversion_list: 
        if not os.path.exists(os.path.join(item[:-1])):
            wb = xl.Workbooks.Open(os.path.join(os.getcwd(), item))
            wb.SaveAs(os.path.join(os.getcwd(), item.replace('.xlsx', '.xls')), FileFormat=56)
            wb.Close()
    xl.Quit()


def process_widetable(file_path):
    wb = xlrd.open_workbook(filename=file_path, formatting_info=True)
    sheet = wb.sheet_by_index(0)

    row_has_number = [False] * sheet.nrows
    for (row, col) in product(range(sheet.nrows), range(sheet.ncols)):
        value = str(sheet.cell(row,col).value)
        if not row_has_number[row] and is_number(value):
            row_has_number[row] = True
    
    col_has_number = [False] * sheet.ncols
    for (row, col) in product(range(sheet.nrows), range(sheet.ncols)):
        value = str(sheet.cell(row, col).value)
        if not col_has_number[col] and is_number(value):
            col_has_number[col] = True
    
    if col_has_number[0]:
        col_head = 1
        col_has_number[0] = False
    else:
        col_head = 0

    col_has_number[col_head] = True
    # set special sheet size
    endrow = 30
    
    output = [["" for x in range(sheet.ncols)] for y in range(endrow)]
    for (row, col) in product(range(endrow), range(sheet.ncols)):
        if is_number(str(sheet.row(row)[col].value)):
            output[row][col] = float(to_number(str(sheet.row(row)[col].value)))
        else:
            output[row][col] = clean_column_name(str(sheet.row(row)[col].value))
    
    output = np.asarray(output)
   
    index_row = 0
    for (row, col) in product(range(len(output)), range(5)):
        if output[row][col] == str(float(col)):
            # print(index_row)
            index_row = row
            break

    for col in range(output.shape[1]):
        if is_number(output[index_row][col]) and int(float(output[index_row][col])) < 0:
            output[index_row][col] = str(abs(int(float(output[index_row][col]))))
            
    # process column header            
    transaction_type = {}
    for col in range(col_head + 1, len(output[0])):
        if col_has_number[col] == True:
            main_cat = []
            for row in range(0, index_row):
                flag = False
                border_index = sheet.cell(row, col).xf_index
                if wb.xf_list[border_index].border.top_line_style != 0:
                    t_type = output[row, col]
                    if wb.xf_list[border_index].border.top_line_style == 6:
                        flag = True
                    if flag:
                        t_type = clean_column_name(t_type)
                        sub_year = str(output[row+1, col])
                        sub_series = clean_column_name(output[row+2, col])
                        if t_type != "":
                            fill_value = t_type
                        else:
                            t_type = fill_value
                        if sub_year != "":
                            fill_year = sub_year
                        else:
                            sub_year = fill_year
                        main_cat.append(t_type)
                        main_cat.append(sub_year)
                        main_cat.append(sub_series)
            if output[index_row, col] != '':   
                transaction_type[int(float(output[index_row, col])) - 1] = main_cat
    
    transactions = pd.DataFrame.from_dict(transaction_type, orient = "index")
    transactions = transactions.reset_index().drop(['index'], axis = 1)
    transactions.columns =['T_type','Year','Series']

    # process outputs
    output = output[:, col_has_number]   
    output[output == ""] = np.nan
    output[output == " "] = np.nan
    output[output == "-"] = np.nan
    output[output == "na"] = np.nan 
        
    output_data = output[index_row + 1:, col_head + 1:]
    
    # get row names here...
    row_names = output[index_row:, 0]
    row_names = row_names[1:,]
    
    DataName = 'IndividualPanel'
    if file_path.endswith('in05st.xls'):
        DataName = 'CrossSectional'
    
    # generate subtable dataframe
    result = []
    for row in range(0, len(output_data)):
        for col in range(0, len(output_data[0])):
            Asset_name = to_camelcase(row_names[row])
            Transaction = to_camelcase(transactions.loc[col,'T_type'])
            year = str(int(float(transactions.loc[col,'Year'])))
            SeriesName =  to_camelcase(transactions.loc[col,'Series'])
            result.append({"Value": output_data[row][col], \
                           "TransactionType":Transaction,\
                           "AssetType":Asset_name, \
                           "Series":  SeriesName,\
                           "Year": year, \
                           "TableID": "Table1ShortTermAndLongTermCapitalGainsAndLosses" ,\
                           "Data": DataName
                                   })
    
    widetables = pd.DataFrame(result).loc[:,["Data", "TableID", "Year",\
                            "TransactionType","AssetType","Series","Value"]]

    return widetables


def process_subtable(file_path, year):
    
    wb = xlrd.open_workbook(filename = file_path, formatting_info = True)
    sheet = wb.sheet_by_index(0)

    if year > 2009:
        subtable_row_start, subtable_row_end, table_name = get_subtables(wb, sheet)    
    elif year in [2007,2008,2009]:  
        subtable_row_start, subtable_row_end, table_name = get_subtables_2(wb, sheet)
    else: 
        subtable_row_start, subtable_row_end, table_name = get_subtables_3(wb, sheet)
    
    row_has_number = [False] * sheet.nrows
    for (row, col) in product(range(sheet.nrows), range(sheet.ncols)):
        value = str(sheet.cell(row,col).value)
        # special diclosure masks
        if value == "**":
            value = str("999999")
        if not row_has_number[row] and is_number(value):
            row_has_number[row] = True
    
    col_has_number = [False] * sheet.ncols
    for (row, col) in product(range(sheet.nrows), range(sheet.ncols)):
        value = str(sheet.cell(row, col).value)
        if not col_has_number[col] and is_number(value):
            col_has_number[col] = True
    
    if col_has_number[0]:
        col_head = 1
        col_has_number[0] = False
    else:
        col_head = 0
    
    col_has_number[col_head] = True
    
    results = pd.DataFrame()
    # process each sub tables
    for i in range(len(subtable_row_start)):
        row_size = subtable_row_end[i] - subtable_row_start[i]
        
        output = [["" for x in range(sheet.ncols)] for y in range(row_size)]
        for (row, col) in product(range(subtable_row_start[i], subtable_row_end[i]), range(sheet.ncols)):
            if is_number(str(sheet.row(row)[col].value)):
                output[row - subtable_row_start[i]][col] = float(to_number(str(sheet.row(row)[col].value)))
            else:
                output[row - subtable_row_start[i]][col] = clean_column_name(str(sheet.row(row)[col].value))
        
        output = np.asarray(output)
        index_row = 0
        for (row, col) in product(range(len(output)), range(6)):
            if output[row][col] == str(col):
                index_row = row
                break
        
        if year > 2000:
            for col in range(output.shape[1]):
                if is_number(output[index_row][col]) and int(output[index_row][col]) < 0:
                    output[index_row][col] = str(abs(int(output[index_row][col])))
            
            # process column header            
            transaction_type = {}
            for col in range(col_head + 1, len(output[0])):
                if col_has_number[col] == True:
                    main_cat = []
                    for row in range(0, index_row):
                        flag = False
                        border_index = sheet.cell(row + subtable_row_start[i], col).xf_index
                        if wb.xf_list[border_index].border.top_line_style != 0:
                            t_type = output[row, col]
                            if year not in [2008,2009]:
                                if wb.xf_list[border_index].border.top_line_style == 6:
                                    flag = True
                            else:
                                border_index_pre =sheet.cell(row + subtable_row_start[i] -1, col).xf_index
                                if wb.xf_list[border_index_pre]._border_flag == 0 and \
                                wb.xf_list[border_index].border.bottom_line_style == 1:
                                    flag = True
                            if flag:
                                t_type = clean_column_name(t_type)
                                sub_series = clean_column_name(output[row+1, col])
                                if t_type != "":
                                    fill_value = t_type
                                else:
                                    t_type = fill_value
                                main_cat.append(t_type)
                                main_cat.append(sub_series)
                        
                    if output[index_row, col] != '':   
                        transaction_type[int(output[index_row, col]) - 1] = main_cat
            
            transactions = pd.DataFrame.from_dict(transaction_type, orient = "index")
            transactions = transactions.reset_index().drop(['index'], axis = 1)
            transactions.columns =['T_type','Series']
            
            # process outputs
            output = output[:, col_has_number]   
            output[output == ""] = np.nan
            output[output == " "] = np.nan
            output[output == "-"] = np.nan
            output[output == "na"] = np.nan 
                
            output_data = output[index_row + 1:, col_head + 1:]
            
            # get row names here...
            row_names = output[index_row:, 0]
            row_names = row_names[1:,]
                        
            # generate subtable dataframe
            result = []
            for row in range(0, len(output_data)):
                for col in range(0, len(output_data[0])):
                    Asset_name = to_camelcase(row_names[row])
                    Transaction = to_camelcase(transactions.loc[col,'T_type'])
                    SeriesName =  to_camelcase(transactions.loc[col,'Series'])
                    result.append({"Value": output_data[row][col], \
                                   "TransactionType":Transaction,\
                                   "AssetType":Asset_name, \
                                   "Series":  SeriesName,\
                                   "Year": year, \
                                   "TableID": table_name[i], \
                                   "Data":'SOIIndividual'
                                           })
            
            subtables = pd.DataFrame(result).loc[:,["Data","TableID", "Year",\
                                    "TransactionType","AssetType","Series","Value"]]
            # combine subtables output
            results = pd.concat([results, subtables])
        
        # for file year 97,98,99
        else:
            latest_number = 1
            latest_col = 0
            unique_row_index_number = set()

            for (row, col) in product(range(row_size), range(1, len(output[0]))):
                if is_number(output[row][col]):
                    #other cases
                    if (is_number(output[row][col - 1]) and float(output[row][col]) - 1 == float(output[row][col - 1]) \
                        and col == latest_col + 1) or\
                    (float(output[row][col]) == latest_number + 1):
                        latest_number += 1
                        latest_col = col
                        unique_row_index_number.add(row)
                        if year == 1998 and row == 51:
                            unique_row_index_number.remove(row)
            
                        
            unique_row_index_number = list(unique_row_index_number)
            unique_row_index_number.sort()
        
            transaction_type = {}
            t_type = ""
            sub_series = ""
            fill_value = ""
            for idx in range(len(unique_row_index_number)):
                index_row = unique_row_index_number[idx]
                for col in range(col_head + 1, len(output[0])):
                    if col_has_number[col] == True:
                        main_cat = []
                        for row in range(index_row-6, index_row):
                            flag = False
                            border_index = sheet.cell(row + subtable_row_start[i], col).xf_index
                            border_index_prev = sheet.cell(row + subtable_row_start[i] -2, col).xf_index
                            # special case for 1999
                            if year == 1999:
                                border_index_prev = sheet.cell(row + subtable_row_start[i] -1, col).xf_index
                                
                            if (wb.xf_list[border_index].border.top_line_style != 0  and idx == 0) or \
                                (wb.xf_list[border_index_prev].border.bottom_line_style !=0  and idx == 1):
                                t_type = output[row, col]
                                if (wb.xf_list[border_index].border.top_line_style == 6 and idx == 0) or \
                                (wb.xf_list[border_index_prev].border.bottom_line_style == 6 and idx == 1) or \
                                (year == 1999 and row == index_row-3 and wb.xf_list[border_index].border.bottom_line_style !=0):
                                    flag = True
                                if flag:
                                    t_type = clean_column_name(t_type)
                                    sub_series = clean_column_name(output[row+1, col])
                                    if t_type != "":
                                        fill_value = t_type
                                    else:
                                        t_type = fill_value
                                    main_cat.append(t_type)
                                    main_cat.append(sub_series)
                            
                        if output[index_row, col] != '':   
                            transaction_type[int(float(output[index_row, col])) - 1] = main_cat
            
            transactions = pd.DataFrame.from_dict(transaction_type, orient = "index")
            transactions = transactions.reset_index().drop(['index'], axis = 1)
            transactions.columns =['T_type','Series']

            if year == 1999:
                for row in range(0, transactions.shape[0]-1):
                    transactions.loc[row, 'T_type'] = transactions.loc[row+1, 'T_type']
                    
            # special column adjustment
            adj_col = 1
            if year == 1997:
                adj_col = 2
                
            for (row, col) in product(range(unique_row_index_number[1],row_size), range(1+adj_col,len(output[0]))): 
                output[row, col-adj_col] = output[row, col]
                output[row, col] = np.nan

            output = output[:, col_has_number]   
            output[output == ""] = np.nan
            output[output == " "] = np.nan
            output[output == "-"] = np.nan 
            output[output == "**"] = np.nan    
            output_data = output[row_has_number[subtable_row_start[i]:subtable_row_end[i]], col_head + 1:]
            
            to_number_fcn = np.vectorize(to_number)
            output_data = to_number_fcn(output_data)
            
            row_names = output[index_row:, 0]
            row_names = row_names[2:,]
            if year == 1997 or year == 1998:
                row_names = row_names[:-1]

            result = []
            #for col in range(1, len(output[0])):
            for row in range(0, len(output_data)):
                if row < len(row_names):   
                    for col in range(0, len(output_data[0])):
                        Asset_name = to_camelcase(row_names[row])
                        Transaction = to_camelcase(transactions.loc[col,'T_type'])
                        SeriesName =  to_camelcase(transactions.loc[col,'Series'])
                        result.append({"Value": output_data[row][col], \
                                       "TransactionType":Transaction,\
                                       "AssetType":Asset_name, \
                                       "Series":  SeriesName,\
                                       "Year": year, \
                                       "TableID": table_name[i], \
                                       "Data":'SOIIndividual'
                                               })
                if row >= len(row_names):
                    for col in range(0, len(output_data[0])-adj_col):
                        Asset_name = to_camelcase(row_names[row-len(row_names)])
                        Transaction = to_camelcase(transactions.loc[col+len(output_data[0]),'T_type'])
                        SeriesName =  to_camelcase(transactions.loc[col+len(output_data[0]),'Series'])
                        result.append({"Value": output_data[row][col], \
                                       "TransactionType":Transaction,\
                                       "AssetType":Asset_name, \
                                       "Series":  SeriesName,\
                                       "Year": year, \
                                       "TableID": table_name[i], \
                                       "Data":'SOIIndividual'
                                               })
            
            subtables = pd.DataFrame(result).loc[:,["Data", "TableID", "Year",\
                                        "TransactionType","AssetType","Series","Value"]]
            
            # combine subtables output
            results = pd.concat([results, subtables])            
        
    return results

def process_soca_table_1():
    
    # load interfaces
    _interface_paths = json.load(open(os.path.join('..','..','.interface_paths.json')))
    # if cache reside with the repository instead of an outside drive.... this is necessary
    os.chdir(os.path.join('..','..'))
    
    if not os.path.exists(os.path.join(_interface_paths["cache"], 'Interfaces')):
        os.makedirs(os.path.join(_interface_paths["cache"], 'Interfaces'))
        
    dir_path = os.path.join(_interface_paths['soca'], 'soca_table_1')
    process_xlsx_conversion(dir_path) 
    
    results = pd.DataFrame()

    file_list_1 = ['07in01ab.xls','08in01soca.xls',\
                 '09in01soca.xls','1001insoca.xls','1101insoca.xls','1201insoca.xls',\
                 '98in5ab.xls','98in1ab.xls','99in01ab.xls']
    # in01st --> individual panel
    # in05st --> cross sectional
    file_list_2 = ['04-07in01st.xls', '99-03in01st.xls','99-03in05st.xls', '04-07in05st.xls']

    
    for filename in file_list_1:
        # for mac error
        if filename != '.DS_Store':
            print ("------")
            print(filename)
            file_path = (os.path.join(_interface_paths['soca'], 'soca_table_1',filename))
            year = filename_year(filename)
            if filename =='98in5ab.xls':
                year = 1997
            result_df = process_subtable(file_path, year)
            results = pd.concat([results, result_df])
            
    for filename in file_list_2:
        if filename != '.DS_Store':
            print ("------")
            print(filename)
            file_path = (os.path.join(_interface_paths['soca'], 'soca_table_1',filename))
            if filename == '04-07in05st.xls':
                file_path = os.path.join('SOI','SOCA', filename)
            result_df = process_widetable(file_path)
            results = pd.concat([results, result_df])
    
    results.to_csv(os.path.join(_interface_paths['cache'],'Interfaces', 'table_1.csv'), index=False)


if __name__ == "__main__":
    process_soca_table_1()