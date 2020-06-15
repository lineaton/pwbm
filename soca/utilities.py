# -*- coding: utf-8 -*-
"""
Created on Tue May  7 16:50:57 2019

@author: Xiaoyue Sun
"""

import re
import xlwt
import os
from itertools import product
from xlutils.copy import copy
import numpy as np
import pandas as pd
import win32com.client

def to_camelcase(val):
	return val \
		.lower() \
		.replace("(", "") \
		.replace(")", "") \
		.replace("  ", " ") \
		.replace(",", "") \
		.replace("-", "") \
		.title() \
		.replace(" ", "") \
		.replace(":", "") \
		.replace(".", "") \
		.replace("\n", "")
        
        
def to_number(val):    
    
    val = val.replace(",", "") \
        .replace("$", "") \
        .replace("*", "") \
        .replace("[", "") \
        .replace("]", "") \
        .replace(" ", "") \
        .replace("--", "") \
        .replace("'", "")\
        .replace("³", "")
    return "0" if val == "" or val == "-" or val == "nan" else val


def is_number(val):
    val = val.replace(",", "") \
        .replace("$", "") \
        .replace("*", "") \
        .replace("[", "") \
        .replace("]", "") \
        .replace(" ", "") \
        .replace("--", "") \
        .replace("'", "")
		#.replace(".", "") \
		
    try:
        float(val)
        return True
    except ValueError:
        return False
    
def clean_column_name(val):
    while len(val) > 0 and val[0] == " ":
        val = val[1:]
    while len(val) > 0 and val[0] == "$":
        val = val[1:]
    val = re.sub("\[(.*?)\]", "", val)
    val = re.sub('<[^>]+>', '', val)

    val =  val.lower() \
        .replace("(", "") \
        .replace(")", "") \
        .replace(" including ", " ") \
        .replace("  ", " ") \
        .replace(",", "") \
        .replace("-", " ") \
        .replace(".", "\n") \
        .replace(":", "") \
        .replace("[", "") \
        .replace("]", "") \
        .replace("\n", " ") \
        .replace("$", " ") \
        .replace(";", "") \
        .replace("continued", "") \
        .replace("&", "and")\
        .replace("/", "")\
        .replace(" cont'd", "")
        
    while len(val) > 0 and val[-1] == " ":
        val = val[:-1]
        
    return "".join(i for i in val if ord(i)<128)

def filename_year(file_name):
    if file_name[0] == '0' or file_name[0] == '1':
        return int("20" + file_name[0:2])
    elif file_name[0] == '9':
        return int("19" + file_name[0:2])


def find_last_row(row_has_number):
    '''find last row of 
    '''
    last_row = len(row_has_number)
    for i in range(75, len(row_has_number) - 5):
        if row_has_number[i] + row_has_number[i + 1] + row_has_number[i + 2] +\
        row_has_number[i + 3] + row_has_number[i + 4] == 0:
            last_row = i
            break
    return last_row


def get_subtables(wb, sheet):
    '''get subtables using the double border index
    '''
    # get subtable start row with double borderd line
    subtable_row_start = []
    for row in range(sheet.nrows-1):
        border_index = sheet.cell(row, 0).xf_index
        if wb.xf_list[border_index].border.bottom_line_style == 6 or \
              wb.xf_list[sheet.cell(row+1, 0).xf_index].border.top_line_style == 6 :
            subtable_row_start.append(row + 1)  
    
    # table end with thin bordered line    
    subtable_row_end = []
    for row in range(sheet.nrows):
        border_index = sheet.cell(row, 1).xf_index
        if (wb.xf_list[border_index].border.top_line_style == 1 and \
            wb.xf_list[border_index].border.right_line_style == 0 and \
             wb.xf_list[border_index].border.left_line_style == 0) or \
            (wb.xf_list[border_index].border.bottom_line_style == 0 and\
            wb.xf_list[border_index].border.left_line_style == 0 and \
            wb.xf_list[border_index].border.right_line_style == 0 and \
            wb.xf_list[sheet.cell(row-1, 1).xf_index].border.bottom_line_style != 0):
            subtable_row_end.append(row)  
    
    table_name = []
    for i in subtable_row_start: 
        tname = sheet.cell((i - 2), 0).value.split(', by Asset Type')[0]
        table_name.append(to_camelcase(tname))
    
    return subtable_row_start, subtable_row_end, table_name


def get_subtables_2(wb, sheet):
    '''get subtables using bold table names
    '''
    # sub table start with bold table name with no borders
    subtable_row_start = []
    for row in range(sheet.nrows):
        border_index = sheet.cell(row, 0).xf_index
        if wb.font_list[wb.xf_list[border_index].font_index].bold == 1 and \
            wb.xf_list[border_index].border.top_line_style == 0 and \
            wb.xf_list[border_index].border.bottom_line_style == 0 and\
            wb.xf_list[border_index].border.left_line_style == 0 and \
            wb.xf_list[border_index].border.right_line_style == 0:
            subtable_row_start.append(row + 2)  
    
    # table ends with the thin border cell 
    subtable_row_end = []
    for row in range(1, sheet.nrows):
        border_index = sheet.cell(row, 0).xf_index
        border_index_prev = sheet.cell(row-1, 0).xf_index
        if wb.xf_list[border_index].border.bottom_line_style == 0 and\
            wb.xf_list[border_index].border.left_line_style == 0 and \
            wb.xf_list[border_index].border.right_line_style == 0 and \
            wb.xf_list[border_index_prev].border.bottom_line_style != 0:
            subtable_row_end.append(row)  
     
    table_name = []
    for i in subtable_row_start: 
        tname = sheet.cell((i - 2), 0).value.split(', by Asset Type')[0]
        table_name.append(to_camelcase(tname))
        
    return subtable_row_start, subtable_row_end, table_name


def get_subtables_3(wb, sheet):        
    subtable_row_start = []
    for row in range(sheet.nrows):
        border_index = sheet.cell(row, 0).xf_index
        if (wb.font_list[wb.xf_list[border_index].font_index].bold == 1) \
        and wb.xf_list[border_index].border.top_line_style == 0 and \
        wb.xf_list[border_index].border.bottom_line_style == 0 and\
        wb.xf_list[border_index].border.left_line_style == 0 and \
        wb.xf_list[border_index].border.right_line_style == 0:
            for ind in range(5):
                new_index = sheet.cell(row+ind, 0).xf_index
                if wb.xf_list[new_index].border.top_line_style == 6:
                    subtable_row_start.append(row)  
                    
    subtable_row_end = []
    for row in range(1,sheet.nrows):
        border_index = sheet.cell(row, 0).xf_index
        if (wb.xf_list[border_index]._border_flag == 0 and \
            wb.xf_list[sheet.cell(row-1, 0).xf_index].border.bottom_line_style == 1):
            subtable_row_end.append(row)  
    
    table_name = []  
    for i in subtable_row_start: 
        tname = sheet.cell(i, 0).value.split(', by Asset Type')[0]
        table_name.append(to_camelcase(tname))
    
    if len(table_name) != 3: 
        check = []
        for i in range(len(table_name)):
            if table_name[i].startswith('Table'):
                check.append(i)
                
        subtable_row_start_new = [subtable_row_start[i] for i in check]
        subtable_row_end_new = [subtable_row_end[i] for i in check]
        table_name_new = [table_name[i] for i in check]
        subtable_row_start = subtable_row_start_new
        subtable_row_end = subtable_row_end_new
        table_name = table_name_new
        
    return subtable_row_start, subtable_row_end, table_name

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