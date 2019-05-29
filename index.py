import os
import win32com.client
import openpyxl
import random
from openpyxl import *

def get_excel_indices(ws, index_headings_col, index_values_col, index_start_row):
    excel_index = {}
    current_row = index_start_row
    while ws[index_headings_col + str(current_row)].value is not None:
        index_heading = ws[index_headings_col + str(current_row)].value
        index_value = ws[index_values_col + str(current_row)].value
        for i in range(len(index_value_array)):
            index_value_array[i] = int(index_value_array[i])
            index_value = index_value_array
        #enter the new entry into the index
        excel_index[index_heading] = index_value
        current_row = current_row + 1
    return excel_index
