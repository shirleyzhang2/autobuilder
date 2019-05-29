import os
import win32com.client
import random
from openpyxl import *

def get_excel_indices(ws, index_headings_col, index_values_col, index_start_row):
    excel_index = {}
    current_row = index_start_row
    while ws[index_headings_col + str(current_row)].value is not None:
        index_heading = ws[index_headings_col + str(current_row)].value
        index_value = ws[index_values_col + str(current_row)].value
        excel_index[index_heading] = index_value
        current_row = current_row + 1
    return excel_index

wb = load_workbook('SetupAB.xlsx')
ws = wb.get_sheet_by_name('Index')
ExcelIndex = Generate_Tower.get_excel_indices(ws, 'A', 'B', 2)

InputTable = ExcelIndex['Input table sheet']
FloorPlan = ExcelIndex['Floor plans sheet']
SectionProperties = ExcelIndex['Section properties sheet']
Bracing = ExcelIndex['Bracing sheet']
FloorBracing = ExcelIndex['Floor bracing sheet']
Materials = ExcelIndex['Materials sheet']
InputTableOffset = ExcelIndex['Input table offset']
PropertiesStartRow = ExcelIndex['Properties start row']
