import os
import win32com.client
import random
from openpyxl import *
from openpyxl.utils.cell import get_column_letter, column_index_from_string
import string

##########get excel index############
def get_excel_indices(wb, index_headings_col, index_values_col, index_start_row):
    excel_index = {}
    ws = wb.get_sheet_by_name('Index')
    current_row = index_start_row
    while ws[index_headings_col + str(current_row)].value is not None:
        index_heading = ws[index_headings_col + str(current_row)].value
        index_value = ws[index_values_col + str(current_row)].value
        #enter the new entry into the index
        excel_index[index_heading] = index_value
        current_row = current_row + 1
    return excel_index

##########read excel tabs############
def get_properties(wb,excel_index,parameter):
    headings_start_col = excel_index['Section or material properties col']
    values_start_col = excel_index['Section or material values col']
    start_row = excel_index['Properties start row']
    if parameter == 'Material':
        ws = wb.get_sheet_by_name('Materials')
    elif parameter == 'Section':
        ws = wb.get_sheet_by_name('Section Properties')
    else:
        print('Input should be either "Material" or"Section"')
    #parameter = 'unknown'
    #if ws['A1'].value == 'Section #':
    #    parameter = 'Section'
    #else:
    #    parameter = 'Material'
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


class Tower:
    def __init__(self, number, member_mat, rod_mat, floor_plans, floor_heights, col_props, bracing_types, floor_masses, floor_bracing_types):
        self.number = number
        self.member_mat = member_mat
        self.rod_mat = rod_mat
        self.floor_plans = floor_plans
        self.floor_heights = floor_heights
        self.col_props = col_props
        self.bracing_types = bracing_types
        self.floor_masses = floor_masses
        self.floor_bracing_types = floor_bracing_types

#Outputs a list containing tower objects representing each tower to be built
def read_input_table(wb,excel_index):
    #Read in the top value in the sheet
    input_table_sheet = excel_index['Input table sheet']
    #Set worksheet to the input table sheet
    ws_input = wb.get_sheet_by_name('Input Table')
    input_table_offset = excel_index['Input table offset']
    total_towers = excel_index['Total number of towers']
    cur_tower_row = 1
    table_start_row = cur_tower_row + input_table_offset
    all_towers = []
    cur_tower_num = 1
    while cur_tower_num <= total_towers:
        #create tower object
        cur_tower = Tower(0,0,0,0,0,0,0,0,0)
        #read member material
        member_mat = ws_input['B'+str(cur_tower_row + 1)].value
        cur_tower.member_mat = member_mat
        #read rod material
        rod_mat = ws_input['B'+str(cur_tower_row + 2)].value
        cur_tower.rod_mat = rod_mat
        #read floor plans
        floor_pln_col = excel_index['Floor plan col']
        cur_floor_row = cur_tower_row + input_table_offset
        floor_plans = []
        while ws_input[floor_pln_col + str(cur_floor_row)].value is not None:
            floor_pln = ws_input[floor_pln_col + str(cur_floor_row)].value
            floor_plans.append(floor_pln)
            cur_floor_row = cur_floor_row + 1
        cur_tower.floor_plans = floor_plans
        #read floor heights
        floor_heights_col = excel_index['Floor height col']
        cur_floor_row = cur_tower_row + input_table_offset
        floor_heights = []
        while ws_input[floor_heights_col + str(cur_floor_row)].value is not None:
            floor_height = ws_input[floor_heights_col + str(cur_floor_row)].value
            floor_heights.append(floor_height)
            cur_floor_row = cur_floor_row + 1
        cur_tower.floor_heights = floor_heights
        #read column properties
        col_props_start_col = excel_index['Column properties start']
        col_props_end_col = excel_index['Column properties end']
        cur_floor_row = cur_tower_row + input_table_offset
        cur_col = col_props_start_col
        col_props = []
        while ws_input[cur_col + str(cur_floor_row)].value is not None:
            while cur_col != chr(ord(col_props_end_col)+1):
                col_prop = ws_input[cur_col + str(cur_floor_row)].value
                col_props.append(col_prop)
                cur_col = chr(ord(cur_col) + 1)
            cur_col = col_props_start_col
            cur_floor_row = cur_floor_row + 1
        cur_tower.col_props = col_props
        #read bracing properties
        bracing_types_start_col = excel_index['Bracing type start']
        bracing_types_end_col= excel_index['Bracing type end']
        cur_floor_row = cur_tower_row + input_table_offset
        cur_col = bracing_types_start_col
        bracing_types = []
        while ws_input[cur_col + str(cur_floor_row)].value is not None:
            while cur_col != chr(ord(bracing_types_end_col)+1):
                bracing_type = ws_input[cur_col + str(cur_floor_row)].value
                bracing_types.append(bracing_type)
                cur_col = chr(ord(cur_col) + 1)
            cur_col = bracing_types_start_col
            cur_floor_row = cur_floor_row + 1
        cur_tower.bracing_types = bracing_types
        #read floor masses
        floor_masses_col = excel_index['Floor mass col']
        cur_floor_row = cur_tower_row + input_table_offset
        floor_masses = []
        while ws_input[floor_masses_col + str(cur_floor_row)].value is not None:
            floor_mass = ws_input[floor_masses_col + str(cur_floor_row)].value
            floor_masses.append(floor_mass)
            cur_floor_row = cur_floor_row + 1
        cur_tower.floor_masses = floor_masses
        #read floor bracing types
        floor_bracing_col = excel_index['Floor bracing col']
        cur_floor_row = cur_tower_row + input_table_offset
        floor_bracing_types = []
        while ws_input[floor_bracing_col + str(cur_floor_row)].value is not None:
            floor_bracing = ws_input[floor_bracing_col + str(cur_floor_row)].value
            floor_bracing_types.append(floor_bracing)
            cur_floor_row = cur_floor_row + 1
        cur_tower.floor_bracing_types = floor_bracing_types
        #increment
        cur_tower.number = cur_tower_num
        all_towers.append(cur_tower)
        cur_tower_num = cur_tower_num + 1
        cur_tower_row = cur_floor_row + 1
    return all_towers

def get_node_info(wb, excel_index,parameter):

    if parameter == 'Floor Bracing':
        ws = wb.get_sheet_by_name('Floor Bracing')
    if parameter == 'Bracing':
        ws = wb.get_sheet_by_name('Bracing')
    if parameter == 'Floor Plans':
        ws = wb.get_sheet_by_name('Floor Plans')

    headings_col = excel_index['Node name col']
    horiz_col = excel_index['Node horiz col']
    vert_col = excel_index['Node vert col']
    start_row = excel_index['Properties start row']

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

def get_floor_or_bracing(wb,excel_index,parameter):

    SectionProperties = get_properties(wb,ExcelIndex,'Section')
    Nodes = get_node_info(wb,ExcelIndex,'Bracing')

    headings_col = excel_index['Floor or bracing name col']
    section_col = excel_index['Floor or bracing section col']
    start_node_col = excel_index['Floor or bracing start node col']
    end_node_col = excel_index['Floor or bracing end node col']
    start_row = excel_index['Properties start row']

    if parameter == 'Floor Bracing':
        ws = wb.get_sheet_by_name('Floor Bracing')
    elif parameter == 'Bracing':
        ws = wb.get_sheet_by_name('Bracing')
    elif parameter == 'Floor Plans':
        ws = wb.get_sheet_by_name('Floor Plans')
    else:
        print('Input should be either "Floor Bracing", "Bracing", or "Floor Plans"')

    #parameter = 'unknown'
    #if ws['A1'].value == 'Bracing #':
    #    parameter = 'Bracing '
    #else:
    #    parameter = 'Floor Plan '

    bracing_index = {}
    current_headings_col = headings_col
    current_section_col = section_col
    current_start_node_col = start_node_col
    current_end_node_col = end_node_col
    i = 1
    while ws[current_headings_col+str(4)].value is not None:
        bracing_index[parameter+' '+str(i)] = {}
        current_row = start_row
        j = 1
        while ws[current_headings_col + str(current_row)].value is not None:
            bracing_index[parameter+' '+str(i)]['Member '+str(j)] = {}
            section = ws[current_section_col + str(current_row)].value
            start_node = ws[current_start_node_col + str(current_row)].value
            end_node = ws[current_end_node_col + str(current_row)].value
            #enter the new entry into the index
            bracing_index[parameter+' '+str(i)] ['Member '+str(j)]['section type']=SectionProperties["Section "+str(section)]
            bracing_index[parameter+' '+str(i)] ['Member '+str(j)]['nodes']=[Nodes["Node "+str(start_node)],Nodes["Node "+str(end_node)]]
            current_row = current_row + 1
            j += 1
        i += 1
        current_headings_col = get_column_letter(column_index_from_string(current_headings_col)+8)
        current_section_col = get_column_letter(column_index_from_string(current_section_col)+8)
        current_start_node_col = get_column_letter(column_index_from_string(current_start_node_col)+8)
        current_end_node_col = get_column_letter(column_index_from_string(current_end_node_col)+8)
    return bracing_index


#TESTING
wb = load_workbook('SetupAB.xlsx')
ExcelIndex = get_excel_indices(wb, 'A', 'B', 2)
#InputTable = ExcelIndex['Input table sheet']
#FloorPlan = ExcelIndex['Floor plans sheet']
#SectionProperties = ExcelIndex['Section properties sheet']
#Bracing = ExcelIndex['Bracing sheet']
#FloorBracing = ExcelIndex['Floor bracing sheet']
#Materials = ExcelIndex['Materials sheet']
#InputTableOffset = ExcelIndex['Input table offset']
#PropertiesStartRow = ExcelIndex['Properties start row']


#SectionProperties = get_properties(wb,ExcelIndex,'Section')
#Materials = get_properties(wb,ExcelIndex,'Material')
Bracing = get_floor_or_bracing(wb,ExcelIndex,'Bracing')
FloorPlans = get_floor_or_bracing(wb,ExcelIndex,'Floor Plans')
FloorBracing = get_floor_or_bracing(wb,ExcelIndex,'Floor Bracing')

#for keys,values in ExcelIndex.items():
#    print(keys)
#    print(values)

#for keys,values in SectionProperties.items():
#    print(keys)
#    print(values)

#for keys,values in Materials.items():
#    print(keys)
#    print(values)

for keys,values in Bracing.items():
    print(keys)
    print(values)

for keys,values in FloorPlans.items():
    print(keys)
    print(values)

AllTowers = read_input_table(wb, ExcelIndex)
for tower in AllTowers:
    print(tower.number)
    print(tower.member_mat)
    print(tower.rod_mat)
    print(tower.floor_plans)
    print(tower.floor_heights)
    print(tower.col_props)
    print(tower.bracing_types)
    print(tower.floor_masses)
    print(tower.floor_bracing_types)
