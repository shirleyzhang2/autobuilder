import os
import win32com.client
import random
from openpyxl import *
from openpyxl.utils.cell import get_column_letter, column_index_from_string

##########get excel index############
def get_excel_indices(ws, index_headings_col, index_values_col, index_start_row):
    excel_index = {}
    current_row = index_start_row
    while ws[index_headings_col + str(current_row)].value is not None:
        index_heading = ws[index_headings_col + str(current_row)].value
        index_value = ws[index_values_col + str(current_row)].value
        #enter the new entry into the index
        excel_index[index_heading] = index_value
        current_row = current_row + 1
    return excel_index

wb = load_workbook('SetupAB.xlsx')
ws_index = wb.get_sheet_by_name('Index')
ExcelIndex = get_excel_indices(ws_index, 'A', 'B', 2)

InputTable = ExcelIndex['Input table sheet']
FloorPlan = ExcelIndex['Floor plans sheet']
SectionProperties = ExcelIndex['Section properties sheet']
Bracing = ExcelIndex['Bracing sheet']
FloorBracing = ExcelIndex['Floor bracing sheet']
Materials = ExcelIndex['Materials sheet']
InputTableOffset = ExcelIndex['Input table offset']
PropertiesStartRow = ExcelIndex['Properties start row']

#testing
#for keys,values in ExcelIndex.items():
#    print(keys)
#    print(values)

##########read excel tabs############
def get_properties(ws,headings_start_col, values_start_col, start_row):
    
    parameter = 'unknown'
    if ws['A1'].value == 'Section #':
        parameter = 'Section'
    else:
        parameter = 'Material'

    parameter_type = {};
    current_property_col = headings_start_col;
    current_value_col = values_start_col;
    i = 1 
    while ws[current_property_col+str(1)].value is not None:
        current_row = start_row
        parameter_type[parameter+' '+str(i)]={}
        while ws[current_property_col + str(current_row)].value is not None:
            properties = {}
            properties_heading = ws[current_property_col + str(current_row)].value
            properties_value = ws[current_value_col + str(current_row)].value
            #enter the new entry into the index
            parameter_type[parameter+' '+str(i)][properties_heading] = properties_value
            current_row = current_row + 1
        i += 1
        current_property_col = get_column_letter(column_index_from_string(current_property_col)+3)
        current_value_col = get_column_letter(column_index_from_string(current_value_col)+3)
    return parameter_type

ws_section = wb.get_sheet_by_name('Section Properties')
SectionProperties = get_properties(ws_section, 'A', 'B',4) #or use ExcelIndex 

ws_materials = wb.get_sheet_by_name('Materials')
Materials = get_properties(ws_materials,'A','B',4)

#testing
#for keys,values in SectionProperties.items():
    #print(keys)
    #print(values)

#testing
#for keys,values in Materials.items():
#    print(keys)
#    print(values)


def get_node_info(ws, headings_col, horiz_col, vert_col, start_row):
    node_index = {}
    current_row = start_row
    while ws[headings_col + str(current_row)].value is not None:
        index_heading = ws[headings_col+ str(current_row)].value
        horiz = ws[horiz_col + str(current_row)].value
        vert = ws[vert_col + str(current_row)].value
        #enter the new entry into the index
        node_index["Node "+str(index_heading)] = [horiz,vert]

        current_row = current_row + 1
    return node_index

ws_nodes = wb.get_sheet_by_name('Bracing')
Nodes = get_node_info(ws_nodes,'A','B','C',4)

def get_floor_or_bracing(ws,headings_col,section_col,start_node_col,end_node_col,start_row):

    parameter = 'unknown'
    if ws['A1'].value == 'Bracing #':
        parameter = 'Bracing '
    else:
        parameter = 'Floor Plan '

    bracing_index = {}
    current_headings_col = headings_col
    current_section_col = section_col
    current_start_node_col = start_node_col
    current_end_node_col = end_node_col
    i = 1

    while ws[current_headings_col+str(4)].value is not None:
        bracing_index[parameter+str(i)] = {}
        current_row = start_row
        j = 1
        while ws[current_headings_col + str(current_row)].value is not None:
            bracing_index[parameter+str(i)]['Member '+str(j)] = {}
            section = ws[current_section_col + str(current_row)].value
            start_node = ws[current_start_node_col + str(current_row)].value
            end_node = ws[current_end_node_col + str(current_row)].value
            #enter the new entry into the index
            bracing_index[parameter+str(i)] ['Member '+str(j)]['section type']=SectionProperties["Section "+str(section)]
            bracing_index[parameter+str(i)] ['Member '+str(j)]['nodes']=[Nodes["Node "+str(start_node)],Nodes["Node "+str(end_node)]]
            current_row = current_row + 1
            j += 1
        i += 1
        current_headings_col = get_column_letter(column_index_from_string(current_headings_col)+8)
        current_section_col = get_column_letter(column_index_from_string(current_section_col)+8)
        current_start_node_col = get_column_letter(column_index_from_string(current_start_node_col)+8)
        current_end_node_col = get_column_letter(column_index_from_string(current_end_node_col)+8)
    return bracing_index

ws_bracing = wb.get_sheet_by_name('Bracing')
ws_floor_plans = wb.get_sheet_by_name('Floor Plans')
Bracing = get_floor_or_bracing(ws_floor_plans,'D','E','F','G',4)

for keys,values in Bracing.items():
    print(keys)
    print(values)

  

#def get_bracing_scheme(ws, node_start_col, section_start_col, start_row):
#    current_row = start_row
#    bracing_scheme_index = {}
#    while ws[headings_col + str(current_row)].value is not none:

