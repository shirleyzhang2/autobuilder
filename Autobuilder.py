import os
import win32com.client
import openpyxl
import random
from openpyxl import *
import re
import time
import ReadExcel

def build_floor_plan_and_bracing(SapModel, tower, all_floor_plans, all_floor_bracing, floor_num, floor_elev):
    print('Building floor plan...')
    floor_plan_num = tower.floor_plans[floor_num-1]
    floor_plan = all_floor_plans[floor_plan_num-1]
    #Create members for floor plan
    for member in floor_plan.members:
        kip_in_F = 3
        SapModel.SetPresentUnits(kip_in_F)
        start_node = member.start_node
        end_node = member.end_node
        start_x = start_node[0]
        start_y = start_node[1]
        start_z = floor_elev
        end_x = end_node[0]
        end_y = end_node[1]
        end_z = floor_elev
        section_name = member.sec_prop
        [ret, name] = SapModel.FrameObj.AddByCoord(start_x, start_y, start_z, end_x, end_y, end_z, PropName=section_name)
        if ret != 0:
            print('ERROR creating floor plan member on floor ' + str(floor_num))
    #assign masses to mass nodes and create steel rod
    mass_node_1 = floor_plan.mass_nodes[0]
    mass_node_2 = floor_plan.mass_nodes[1]
    floor_mass = tower.floor_masses[floor_num-1]
    mass_per_node = floor_mass/2
    #Create the mass node point
    [ret, mass_name_1] = SapModel.PointObj.AddCartesian(mass_node_1[0],mass_node_1[1],floor_elev,MergeOff=False)
    if ret != 0:
        print('ERROR setting mass nodes on floor ' + str(floor_num))
    [ret, mass_name_2] = SapModel.PointObj.AddCartesian(mass_node_2[0],mass_node_2[1],floor_elev,MergeOff=False)
    if ret != 0:
        print('ERROR setting mass nodes on floor ' + str(floor_num))
    #Assign masses to the mass nodes
    #Shaking in the x direcion!
    N_m_C = 10
    SapModel.SetPresentUnits(N_m_C)
    ret = SapModel.PointObj.SetMass(mass_name_1, [mass_per_node,0,0,0,0,0],0,True,False)
    if ret[0] != 0:
        print('ERROR setting mass on floor ' + str(floor_num))
    ret = SapModel.PointObj.SetMass(mass_name_2, [mass_per_node,0,0,0,0,0])
    if ret[0] != 0:
        print('ERROR setting mass on floor ' + str(floor_num))
    #Create steel rod
    kip_in_F = 3
    SapModel.SetPresentUnits(kip_in_F)
    [ret, name1] = SapModel.FrameObj.AddByCoord(mass_node_1[0], mass_node_1[1], floor_elev, mass_node_2[0], mass_node_2[1], floor_elev, PropName='Steel rod')
    if ret !=0:
        print('ERROR creating steel rod on floor ' + str(floor_num))
    #Create floor load forces
    N_m_C = 10
    SapModel.SetPresentUnits(N_m_C)
    ret = SapModel.PointObj.SetLoadForce(mass_name_1, 'DEAD', [0, 0, mass_per_node*9.81, 0, 0, 0])
    ret = SapModel.PointObj.SetLoadForce(mass_name_2, 'DEAD', [0, 0, mass_per_node*9.81, 0, 0, 0])
    #create floor bracing
    floor_bracing_num = tower.floor_plans[floor_num-1]
    floor_bracing = all_floor_bracing[floor_bracing_num-1]
    #Finding x and y scaling factors:
    all_plan_nodes = []
    for member in floor_plan.members:
        all_plan_nodes.append(member.start_node)
        all_plan_nodes.append(member.end_node)
    #Find max and min x and y coordinates
    max_node_x = 0
    max_node_y = 0
    min_node_x = 0
    min_node_y = 0
    for node in all_plan_nodes:
        if max_node_x < node[0]:
            max_node_x = node[0]
        if max_node_y < node[1]:
            max_node_y = node[1]
        if min_node_x > node[0]:
            min_node_x = node[0]
        if min_node_y > node[1]:
            min_node_y = node[1]
    scaling_x = max_node_x - min_node_x
    scaling_y = max_node_y - min_node_y
    #Create floor bracing
    print('Building floor bracing...')
    for member in floor_bracing.members:
        kip_in_F = 3
        SapModel.SetPresentUnits(kip_in_F)
        start_node = member.start_node
        end_node = member.end_node
        start_x = start_node[0] * scaling_x
        start_y = start_node[1] * scaling_y
        start_z = floor_elev
        end_x = end_node[0] * scaling_x
        end_y = end_node[1] * scaling_y
        end_z = floor_elev
        section_name = member.sec_prop
        [ret, name] = SapModel.FrameObj.AddByCoord(start_x, start_y, start_z, end_x, end_y, end_z, PropName=section_name)
        if ret != 0:
            print('ERROR creating floor bracing member on floor ' + str(floor_num))
    return SapModel, scaling_x, scaling_y

def build_face_bracing(SapModel, tower, all_floor_plans, all_face_bracing, floor_num, floor_elev):
    print('Building face bracing...')
    i = 1
    while i <= len(Tower.side):
        face_bracing_num = Tower.bracing_types[floor_num][i-1]
        face_bracing = all_face_bracing[face_bracing_num-1]

        #Find max x and y coordinates:
        all_plan_nodes = []
        floor_plan_num = tower.floor_plans[floor_num-1]
        floor_plan = all_floor_plans[floor_plan_num-1]
        for member in floor_plan.members:
            all_plan_nodes.append(member.start_node)
            all_plan_nodes.append(member.end_node)

        all_bracing_nodes = []
        for member in face_bracing.members:
            all_bracing_nodes.append(member.start_node)
            all_bracing_nodes.append(member.end_node)

        max_node_x = 0
        max_node_y = 0
        min_node_x = 0
        min_node_y = 0
        for node in all_plan_nodes:
            if max_node_x < node[0]:
                max_node_x = node[0]
            if max_node_y < node[1]:
                max_node_y = node[1]
            if min_node_x > node[0]:
                min_node_x = node[0]
            if min_node_y > node[1]:
                min_node_y = node[1]
        #Find scaling factor
        scaling_x = max_node_x - min_node_x
        scaling_y = max_node_y - min_node_y
        scaling_z = tower.floor_heights[floor_num-1]
        
        for member in face_bracing.members:
            kip_in_F = 3
            SapModel.SetPresentUnits(kip_in_F)
            start_node = member.start_node
            end_node = member.end_node
            
            #Create face bracing for long side
            if i == ExcelIndex['Side 1'] or i == ExcelIndex['Side 3']:
                start_x = start_node[0] * scaling_x
                start_y = 0
                start_z = start_node[1] * scaling_z + floor_elev
                end_x = end_node[0] * scaling_x
                end_y = 0
                end_z = end_node[1] * scaling_z + floor_elev
            #Create face bracing for short side
            elif i == ExcelIndex['Side 2'] or i == ExcelIndex['Side 4']:
                start_x = start_node[0] * scaling_y
                start_y = 0
                start_z = start_node[1] * scaling_z + floor_elev
                end_x = end_node[0] * scaling_y
                end_y = 0 
                end_z = end_node[1] * scaling_z + floor_elev
            section_name = member.sec_prop 
            #rotate coordinate system 

            if i == ExcelIndex['Side 1']:
                ret = SapModel.CoordSys.SetCoordSys('CSys1', 0, 0, 0, 0, 0, 0)
            elif i == ExcelIndex['Side 2']:
                ret = SapModel.CoordSys.SetCoordSys('CSys1', scaling_x, 0, 0, 90, 0, 0)
            elif i == ExcelIndex['Side 3']:
                ret = SapModel.CoordSys.SetCoordSys('CSys1', 0, scaling_y, 0, 0, 0, 0)
            elif i == ExcelIndex['Side 4']:
                ret = SapModel.CoordSys.SetCoordSys('CSys1', 0, 0, 0, 90, 0, 0)

            [ret, name] = SapModel.FrameObj.AddByCoord(start_x, start_y, start_z, end_x, end_y, end_z, ' ', section_name, ' ', 'CSys1')
            if ret != 0:
                print('ERROR creating floor bracing member on floor ' + str(floor_num))
        #change coordinate system depending on long/short side
        #if i/2 != 1.0:
            #ret = SapModel.CoordSys.SetCoordSys('CSys1', scaling_x, 0, 0, 90, 0, 0)
        #else:
            #ret = SapModel.CoordSys.SetCoordSys('CSys1', 0, scaling_y, 0, 90, 0, 0)
        i += 1
    return SapModel

def get_acc_and_drift(SapObject):
    #Run Analysis
    print('Computing accelaration and drift...')
    SapModel.Analyze.RunAnalysis()
    print('Finished computing.')
    #Get RELATIVE acceleration from node
    SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
    SapModel.Results.Setup.SetComboSelectedForOutput('DEAD + GM', True)
    #set type to envelope
    SapModel.Results.Setup.SetOptionModalHist(1)
    #Get joint acceleration
    #Set units to metres
    N_m_C = 10
    SapModel.SetPresentUnits(N_m_C)
    g = 9.81
    ret = SapModel.Results.JointAccAbs('0-0-0', 0)#for now
    max_and_min_acc = ret[7]
    max_pos_acc = max_and_min_acc[0]
    min_neg_acc = max_and_min_acc[1]
    if abs(max_pos_acc) >= abs(min_neg_acc):
        max_acc = abs(max_pos_acc)/g
    elif abs(min_neg_acc) >= abs(max_pos_acc):
        max_acc = abs(min_neg_acc)/g
    else:
        print('Could not find max acceleration')
    #Get joint displacement
    #Set units to millimetres
    N_mm_C = 9
    SapModel.SetPresentUnits(N_mm_C)
    ret = SapModel.Results.JointDispl('0-0-0', 0)#for now
    max_and_min_disp = ret[7]
    max_pos_disp = max_and_min_disp[0]
    min_neg_disp = max_and_min_disp[1]
    if abs(max_pos_disp) >= abs(min_neg_disp):
        max_drift = abs(max_pos_acc)
    elif abs(min_neg_disp) >= abs(max_pos_disp):
        max_drift = abs(min_neg_disp)
    else:
        print('Could not find max drift')
    #Close SAP2000
    SapObject.ApplicationExit(True)
    return max_acc, max_drift

def print_acc_and_drift(SapObject):
    print('\nAnalyze')
    print('----------------------------------')
    max_acc_and_drift = get_sap_results(SapObject)
    print('Max acceleration is: ' + str(max_acc_and_drift[0]) + ' g')
    print('Max drift is: ' + str(max_acc_and_drift[1]) + ' mm')
    return max_acc_and_drift

def get_weight(SapObject):
    #Run Analysis
    print('Computing weight...')
    SapModel.Analyze.RunAnalysis()
    print('Finished computing.')
    #Get base reactions
    SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
    SapModel.Results.Setup.SetCaseSelectedForOutput('DEAD')
    #SapModel.Results.BaseReact(NumberResults, LoadCase, StepType, StepNum, Fx, Fy, Fz, Mx, My, Mz, gx, gy, gz)
    ret = SapModel.Results.BaseReact()
    base_react = ret[7]
    return base_react

def get_FABI(SAPObject):
    results = get_acc_and_drift(SapObject)
    footprint = 96 #inches squared
    weight = get_weight(SapObject) #lb
    design_life = 100 #years
    construction_cost = 2500000*(weight**2)+6*(10**6)
    land_cost = 35000 * footprint
    annual_building_cost = (land_cost + construction_cost) / design_life
    annual_revenue = 430300
    equipment_cost = 20000000
    return_period_1 = 50
    return_period_2 = 300
    max_disp = results[1] #mm
    apeak_1 = results[0] #g's
    xpeak_1 = 100*max_disp/1524 #% roof drift
    structural_damage_1 = scipy.stats.norm(1.5, 0.5).cdf(xpeak_1)
    equipment_damage_1 = scipy.stats.norm(1.75, 0.7).cdf(apeak_1)
    economic_loss_1 = structural_damage_1*construction_cost + equipment_damage_1*equipment_cost
    annual_economic_loss_1 = economic_loss_1/return_period_1
    structural_damage_2 = 0.5
    equipment_damage_2 = 0.5
    economic_loss_2 = structural_damage_2*construction_cost + equipment_damage_2*equipment_cost
    annual_economic_loss_2 = economic_loss_2/return_period_2
    annual_seismic_cost = annual_economic_loss_1 + annual_economic_loss_2
    fabi = annual_revenue - annual_building_cost - annual_seismic_cost
    return fabi

def write_to_excel(SapObject):
    wb = openpyxl.Workbook()
    ws = wb.active
    #ws = wb.create_sheet(title = "FABI")
    ws['A1'] = 'Tower #'
    ws['A2'] = 'FABI'
    for tower in AllTowers:
        col = get_column_letter(tower.number+1)
        ws[col +'1'] = tower.number
        ws[col +'2'] = fabi
    wb.save('C:\\Users\\shirl\\OneDrive - University of Toronto\\Desktop\\Seismic\\FABI.xlsx')



#----START-----------------------------------------------------START----------------------------------------------------#



print('\n--------------------------------------------------------')
print('Autobuilder by University of Toronto Seismic Design Team')
print('--------------------------------------------------------\n')

#Read in the excel workbook
print("\nReading Excel spreadsheet...")
wb = load_workbook('SetupAB.xlsm')
ExcelIndex = ReadExcel.get_excel_indices(wb, 'A', 'B', 2)

Sections = ReadExcel.get_properties(wb,ExcelIndex,'Section')
Materials = ReadExcel.get_properties(wb,ExcelIndex,'Material')
Bracing = ReadExcel.get_bracing(wb,ExcelIndex,'Bracing')
FloorPlans = ReadExcel.get_floor_plans(wb,ExcelIndex)
FloorBracing = ReadExcel.get_bracing(wb,ExcelIndex,'Floor Bracing')
AllTowers = ReadExcel.read_input_table(wb, ExcelIndex)

print('\nInitializing SAP2000 model...')
# create SAP2000 object
SapObject = win32com.client.Dispatch('SAP2000v15.SapObject')
# start SAP2000
SapObject.ApplicationStart()
# create SapModel Object
SapModel = SapObject.SapModel
# initialize model
SapModel.InitializeNewModel()
# create new blank model
ret = SapModel.File.NewBlank()

#Define new materials
print("\nDefining materials...")
N_m_C = 10
SapModel.SetPresentUnits(N_m_C)
for Material, MatProps in Materials.items():
    MatName = MatProps['Name']
    MatType = MatProps['Material type']
    MatWeight = MatProps['Weight per volume']
    MatE = MatProps['Elastic modulus']
    MatPois = MatProps['Poisson\'s ratio']
    MatTherm = MatProps['Thermal coefficient']
    #Create material type
    ret = SapModel.PropMaterial.SetMaterial(MatName, MatType)
    if ret != 0:
        print('ERROR creating material type')
    #Set isotropic material proprties
    ret = SapModel.PropMaterial.SetMPIsotropic(MatName, MatE, MatPois, MatTherm)
    if ret != 0:
        print('ERROR setting material properties')
    #Set unit weight
    ret = SapModel.PropMaterial.SetWeightAndMass(MatName, 1, MatWeight)
    if ret != 0:
        print('ERROR setting material unit weight')

#Define new sections
print('Defining sections...')
kip_in_F = 3
SapModel.SetPresentUnits(kip_in_F)
for Section, SecProps in Sections.items():
    SecName = SecProps['Name']
    SecArea = SecProps['Area']
    SecTors = SecProps['Torsional constant']
    SecIn3 = SecProps['Moment of inertia about 3 axis']
    SecIn2 = SecProps['Moment of inertia about 2 axis']
    SecSh2 = SecProps['Shear area in 2 direction']
    SecSh3 = SecProps['Shear area in 3 direction']
    SecMod3 = SecProps['Section modulus about 3 axis']
    SecMod2 = SecProps['Section modulus about 2 axis']
    SecPlMod3 = SecProps['Plastic modulus about 3 axis']
    SecPlMod2 = SecProps['Plastic modulus about 2 axis']
    SecRadGy3 = SecProps['Radius of gyration about 3 axis']
    SecRadGy2 = SecProps['Radius of gyration about 2 axis']
    SecMat = SecProps['Material']
    #Create section property
    ret = SapModel.PropFrame.SetGeneral(SecName, SecMat, 0.1, 0.1, SecArea, SecSh2, SecSh3, SecTors, SecIn2, SecIn3, SecMod2, SecMod3, SecPlMod2, SecPlMod3, SecRadGy2, SecRadGy3, -1)
    if ret != 0:
        print('ERROR creating section property ' + SecName)

TowerNum = 1
for Tower in AllTowers:
    print('\nBuilding tower number ' + str(TowerNum))
    print('-------------------------')
    print(Tower.bracing_types)
    NumFloors = len(Tower.floor_plans)
    CurFloorNum = 1
    CurFloorElevation = 0
    while CurFloorNum <=  NumFloors:
        print('Floor ' + str(CurFloorNum))
        if CurFloorNum <=  NumFloors:
            build_floor_plan_and_bracing(SapModel, Tower, FloorPlans, FloorBracing, CurFloorNum, CurFloorElevation)
        if CurFloorNum <  NumFloors:
            build_face_bracing(SapModel, Tower, FloorPlans, Bracing, CurFloorNum, CurFloorElevation)
        #INSERT FUNCTION TO CREATE COLUMNS AT CURRENT FLOOR

        CurFloorHeight = Tower.floor_heights[CurFloorNum - 1]
        CurFloorElevation = CurFloorElevation + CurFloorHeight
        CurFloorNum += 1

    TowerNum += 1
