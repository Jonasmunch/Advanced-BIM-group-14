###############################################
#### Initial code info                     ####
###############################################

# To run code following 3 packages must be installed to pip (pip install [package]): 
# (1) ExcelWriter  (2) numpy



###############################################
#### Model load                            ####
###############################################

import tkinter as tk
from tkinter.filedialog import askopenfilename
tk.Tk().withdraw() # part of the import if you are not using other tkinter functions

Ifc_path = askopenfilename() #The path the the Ifc file is saved at Ifc_path

import ifcopenshell
file = ifcopenshell.open(Ifc_path)  #The selected Ifc file gets imported.

import xlsxwriter
import numpy

###############################################
#### Element list Extraction               ####
###############################################

# Creating element vectors for later property extraction.
beams = file.by_type('IfcBeam')
columns = file.by_type('IfcColumn') 
Walls = file.by_type('IfcWall')
slabs = file.by_type('IfcSlab')

# Creating empty vectors, for easy element numbering.
beam_num = []
column_num = []
wall_num = []
slab_num = []


###############################################
#### Entity Extraction                     ####
###############################################

######## BEAMS:
# NOTE: 
  # Expected information vectors named with associated entity:
    # beam_El_name,  beam_Material,  beam_Lenght,  beam_placement,  beam_crossection
num = 0

beam_El_name = []
beam_Material = []
beam_Length = []
beam_placement_x = []
beam_placement_y = []
beam_placement_z = []
beam_level = []
beam_crossection_bf = []
beam_crossection_d = []
beam_crossection_k = []
beam_crossection_kr = []
beam_crossection_tf = []
beam_crossection_tw = []

for beam in beams:
#
    num = num+1
    b_num = "beam"+str(num)
    beam_num.append(b_num)

    #Extracting beam_El_name
    beam_El_name.append(beam.Name)

    #Extracting Reference level
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Constraints":
                call_a = "ERROR"
                for property in property_set.HasProperties:
                    if property.Name == "Reference Level":
                        call_a = property.NominalValue.wrappedValue
                  # Assigning result of logical test's:
                beam_level.append(call_a)

    #Extracting Length
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "Length":
                        call_a = round(property.NominalValue.wrappedValue, 2)
                  # Assigning result of logical test's:
                beam_Length.append(call_a)

    #Extracting Material
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Materials and Finishes":
                for property in property_set.HasProperties:
                    if property.Name == "Beam Material":
                        call_a = property.NominalValue.wrappedValue
                  # Assigning result of logical test's:
                beam_Material.append(call_a)

    #Extracting Crossection
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                call_a = "ERROR" ; call_b = "ERROR" ; call_c = "ERROR" ; call_d = "ERROR" ; call_e = "ERROR" ; call_a = "ERROR"
                for property in property_set.HasProperties:
                    # bf
                    if property.Name == "bf":
                        call_a = round(property.NominalValue.wrappedValue, 2)
                    # d
                    if property.Name == "d":
                        call_b = round(property.NominalValue.wrappedValue, 2)
                    # k
                    if property.Name == "k":
                        call_c = round(property.NominalValue.wrappedValue, 2)
                    # kr
                    if property.Name == "kr":
                        call_d = round(property.NominalValue.wrappedValue, 2)
                    # tf
                    if property.Name == "tf":
                        call_e = round(property.NominalValue.wrappedValue, 2)
                    # tw
                    if property.Name == "tw":
                        call_f = round(property.NominalValue.wrappedValue, 2)
                # Assigning result of logical test's:
                beam_crossection_bf.append(call_a)
                beam_crossection_d.append(call_b)
                beam_crossection_k.append(call_c)
                beam_crossection_kr.append(call_d)
                beam_crossection_tf.append(call_e)
                beam_crossection_tw.append(call_f)

    #Extracting Object coordinates
    xval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[0],3)
    beam_placement_x.append(xval)
    yval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[1],3)
    beam_placement_y.append(yval)
    zval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[2],3)
    beam_placement_z.append(zval)

dataframe_beams = [beam_num, beam_El_name, beam_level, beam_Material, beam_Length, beam_crossection_bf, beam_crossection_d, beam_crossection_k, beam_crossection_kr, beam_crossection_tf, beam_crossection_tw, beam_placement_x, beam_placement_y, beam_placement_z]
beams_dataframe = numpy.transpose(dataframe_beams)


######## COLUMNS:
# NOTE: 
  # Expected information vectors named with associated entity:
    # column_El_name,  column_Material,  column_Lenght,  column_placement,  column_crossection
num = 0

column_El_name = []
column_Material = []
column_Length = []
column_placement_x = []
column_placement_y = []
column_placement_z = []
column_level = []
column_crossection_b = []
column_crossection_h = []
column_crossection_k = []
column_crossection_kr = []
column_crossection_tf = []
column_crossection_tw = []

for column in columns:
#
    num = num+1
    c_num = "column"+str(num)
    column_num.append(c_num)

    #Extracting Reference level
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Constraints":
                call_a = "ERROR"
                for property in property_set.HasProperties:
                    if property.Name == "Reference Level":
                        call_a = property.NominalValue.wrappedValue
                  # Assigning result of logical test's:
                column_level.append(call_a)

    #Extracting Length
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "Length":
                        call_a = round(property.NominalValue.wrappedValue, 2)
                  # Assigning result of logical test's:
                column_Length.append(call_a)

    #Extracting Material
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Materials and Finishes":
                for property in property_set.HasProperties:
                    if property.Name == "Beam Material":
                        call_a = property.NominalValue.wrappedValue
                  # Assigning result of logical test's:
                column_Material.append(call_a)

    #Extracting Crossection
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                call_a = "ERROR" ; call_b = "ERROR" ; call_c = "ERROR" ; call_d = "ERROR" ; call_e = "ERROR" ; call_a = "ERROR"
                for property in property_set.HasProperties:
                    # b
                    if property.Name == "b":
                        call_a = round(property.NominalValue.wrappedValue, 2)
                    # h
                    if property.Name == "h":
                        call_b = round(property.NominalValue.wrappedValue, 2)
                    # k
                    if property.Name == "k":
                        call_c = round(property.NominalValue.wrappedValue, 2)
                    # kr
                    if property.Name == "kr":
                        call_d = round(property.NominalValue.wrappedValue, 2)
                    # tf
                    if property.Name == "tf":
                        call_e = round(property.NominalValue.wrappedValue, 2)
                    # tw
                    if property.Name == "tw":
                        call_f = round(property.NominalValue.wrappedValue, 2)
                # Assigning result of logical test's:
                column_crossection_b.append(call_a)
                column_crossection_h.append(call_b)
                column_crossection_k.append(call_c)
                column_crossection_kr.append(call_d)
                column_crossection_tf.append(call_e)
                column_crossection_tw.append(call_f)


    #Extracting Object coordinates
    xval = round(column.ObjectPlacement.RelativePlacement.Location.Coordinates[0],3)
    column_placement_x.append(xval)
    yval = round(column.ObjectPlacement.RelativePlacement.Location.Coordinates[1],3)
    column_placement_y.append(yval)
    zval = round(column.ObjectPlacement.RelativePlacement.Location.Coordinates[2],3)
    column_placement_z.append(zval)

dataframe_columns = [column_num, column_El_name, column_level, column_Material, column_Length, column_crossection_b, column_crossection_h, column_crossection_k, column_crossection_kr, column_crossection_tf, column_crossection_tw, column_placement_x, column_placement_y, column_placement_z]
columns_dataframe = numpy.transpose(dataframe_columns)



######## WALLS:
# Creating empty property vectors:
wall_num = [] ; wall_El_name = [] ; wall_Height =[] ; wall_Lenght = [] ; wall_Width =[] ; wall_Volume =[] ; wall_Type = []

num = 0


for wall in Walls:
#
    num = num+1
    w_num = "wall"+str(num)
    wall_num.append(w_num)
    wall_El_name.append(wall.Name)
# 
    for definition in wall.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            #
            #Extracting Height of walls
            if property_set.Name == "PSet_Revit_Constraints":
                for property in property_set.HasProperties:
                    if property.Name == "Unconnected Height":
                        UH_w = (round(property.NominalValue.wrappedValue, 2))
                        wall_Height.append(UH_w)
            #
            #Extracting length of wall
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "Length":
                        length_w = (round(property.NominalValue.wrappedValue, 2))
                        wall_Lenght.append(length_w)
            #
            #Extracting thinkness of wall
            if property_set.Name == "PSet_Revit_Type_Construction":
                for property in property_set.HasProperties:
                    if property.Name == "Width":
                        thickness_w = (round(property.NominalValue.wrappedValue, 2))
                        wall_Width.append(thickness_w)
            #
            #Extracting volume of wall
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "Volume":
                        volume_w = (round(property.NominalValue.wrappedValue, 2))
                        wall_Volume.append(volume_w)


    #Extracting wall type (Virker)
    for relAssociatesMaterial in wall.HasAssociations:
        material_wall = relAssociatesMaterial.RelatingMaterial.ForLayerSet.LayerSetName

        material_w = material_wall
        material_w = float('nan')                   # laver NaN af data der ikke er et tal
        wall_type_ch = math.isnan(volume_w)         # tjekker om den fundet værdi er et NaN

        if wall_type_ch == True :                   # Hvis findet værdi er et NaN SÅ:
            material_w_ch = material_wall
        else:                                       # Hvis findet værdi ikke er et NaN SÅ:
            material_w_ch = 'Error'                 

        wall_Type.append(material_w_ch)





######## SLABS:
# Creating empty property vectors:
slabs_structural = [] ; slab_num = [] ; slab_El_name = [] ; slab_Perimeter = [] ; slab_Type = []
slab_Area = [] ; slab_Volume = [] ; slab_Thickness = [] ; slab_Level = []
slab_xval = [] ; slab_yval = [] ; slab_zval = [] ; xval = [] ; yval = [] ; zval = []
num = 0


# Extracting the bearing slabs from floor finishes and exterior slabs:
for slab in slabs:
    for RAM in slab.HasAssociations:
        if "Finish" in RAM.RelatingMaterial.ForLayerSet.LayerSetName:
            continue
        if "Exterior" in RAM.RelatingMaterial.ForLayerSet.LayerSetName:
            continue
        slabs_structural.append(slab)

# Extracting properties:
for slab in slabs_structural:
    num = num+1
    slab_num.append("Slab"+str(num))
    slab_El_name.append(slab.Name)
    #
    # Extracting dimensional properties:
    for definition in slab.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Dimensions":
                call_a = "ERROR"  ;  call_b = "ERROR"  ;  call_c = "ERROR" ; call_d = "ERROR"
                for property in property_set.HasProperties:
                    # Perimeter properties:
                    if property.Name == "Perimeter":
                        call_a = round(property.NominalValue.wrappedValue, 2)
                    # Area Properties:
                    elif property.Name == "Area":
                        call_b = round(property.NominalValue.wrappedValue, 2)
                    # Volume Properties:
                    elif property.Name == "Volume":
                        call_c = round(property.NominalValue.wrappedValue, 2)
                    # Thickness properties:
                    elif property.Name == "Thickness":
                        call_d = round(property.NominalValue.wrappedValue, 2)
                # Assigning result of logical test's:
                slab_Perimeter.append(call_a)
                slab_Area.append(call_b)
                slab_Volume.append(call_c)
                slab_Thickness.append(call_d)
            #
            # Extracting reference level:
            if property_set.Name == "PSet_Revit_Constraints":
                call_a = "ERROR"
                for property in property_set.HasProperties:
                    if property.Name == "Level":
                        call_a = property.NominalValue.wrappedValue
                slab_Level.append(call_a)
    #
    # Extracting Material:
    for RAM in slab.HasAssociations:
        call_a = "ERROR"
        material = RAM.RelatingMaterial.ForLayerSet.LayerSetName
        if material == "":
            slab_Type.append(call_a)
        else:
            slab_Material.append(material)
    #
    # Extracting coordinates
    call_a = "ERROR"
    xval = round(slab.ObjectPlacement.RelativePlacement.Location.Coordinates[0],3)
    yval = round(slab.ObjectPlacement.RelativePlacement.Location.Coordinates[1],3)
    zval = round(slab.ObjectPlacement.RelativePlacement.Location.Coordinates[2],3)
    # searching for errors:
    if xval == "":
        slab_xval.append(call_a)
    else:
        slab_xval.append(xval)
    if yval == "":
        slab_yval.append(call_a)
    else:
        slab_yval.append(yval)
    if zval == "":
        slab_zval.append(call_a)
    else:
        slab_zval.append(zval)


dataframe_slabs = [slab_num, slab_El_name, slab_Level, slab_Type, slab_Thickness, slab_Area, slab_Perimeter, slab_Volume, slab_xval, slab_yval, slab_zval]
slabs_dataframe = numpy.transpose(dataframe_slabs)


###############################################
#### Creating Excel document               ####
###############################################

############
# Naming output file

#import tkinter as tk
from tkinter import simpledialog

ROOT = tk.Tk()
ROOT.withdraw()
# Opening nameing dialog window:
Output_FileN = simpledialog.askstring(title="Test", prompt="Filename of outputfile")   
# Saving custom name til constant:
Output_Filename = str(Output_FileN)+'.xlsx'


############
#Creating workbook and worksheets:
workbook = xlsxwriter.Workbook('Slab_info.xlsx')
worksheet1 = workbook.add_worksheet('Info')
worksheet2 = workbook.add_worksheet('Slabs')
worksheet3 = workbook.add_worksheet('Walls')
worksheet4 = workbook.add_worksheet('Beams')
worksheet5 = workbook.add_worksheet('Columns')


######
# Formatting:
red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})


######
# Creating tables within the sheets:

# SLABS:
rows = len(slabs_structural)
columns = len(dataframe_slabs)-1

worksheet2.add_table(0,0, rows, columns, { 
    'data': slabs_dataframe,
    'style': 'Table Style Light 11',
    'name': "Slabs_data",
    'columns': [{'header': 'Element'}, {'header': 'Element Name'}, {'header': 'Ref_level'}, {'header': 'Type'}, 
    {'header': 'Thickness [m]'}, {'header': 'Area [m^2]'}, {'header': 'Perimeter [m]'}, {'header': 'Volume [m^3]'}, {'header': 'Placement_x-val'},
    {'header': 'Placement_y-val'}, {'header': 'Placement_z-val'},]
})

# WALLS:
worksheet3.add_table(0,0,rows,10, { 
    'data': walls_dataframe,
    'style': 'Table Style Light 11',
    'name': "Walls_data",
    'columns': [{'header': 'Element'}, {'header': 'Element Name'}, {'header': 'Ref_level'}, {'header': 'Material'}, 
    {'header': 'Thickness'}, {'header': 'Area'}, {'header': 'Perimeter'}, {'header': 'Volume'}, {'header': 'Placement_x-val'},
    {'header': 'Placement_y-val'}, {'header': 'Placement_z-val'},]
})

# BEAMS:
worksheet4.add_table(0,0,rows,10, { 
    'data': beams_dataframe,
    'style': 'Table Style Light 11',
    'name': "Beams_data",
    'columns': [{'header': 'Element'}, {'header': 'Element Name'}, {'header': 'Ref_level'}, {'header': 'Material'}, 
    {'header': 'Thickness'}, {'header': 'Area'}, {'header': 'Perimeter'}, {'header': 'Volume'}, {'header': 'Placement_x-val'},
    {'header': 'Placement_y-val'}, {'header': 'Placement_z-val'},]
})

# COLUMNS:
worksheet5.add_table(0,0,rows,10, { 
    'data': columns_dataframe,
    'style': 'Table Style Light 11',
    'name': "Columns_data",
    'columns': [{'header': 'Element'}, {'header': 'Element Name'}, {'header': 'Ref_level'}, {'header': 'Material'}, 
    {'header': 'Thickness'}, {'header': 'Area'}, {'header': 'Perimeter'}, {'header': 'Volume'}, {'header': 'Placement_x-val'},
    {'header': 'Placement_y-val'}, {'header': 'Placement_z-val'},]
})

######
# Conditional formatting: 
worksheet2.conditional_format(slabs_dataframe, {'type': 'cell', 'criteria': 'equal to', 'value': 'ERROR', 'format': 'red_format'})
worksheet3.conditional_format(walls_dataframe, {'type': 'cell', 'criteria': 'equal to', 'value': 'ERROR', 'format': 'red_format'})
worksheet4.conditional_format(beams_dataframe, {'type': 'cell', 'criteria': 'equal to', 'value': 'ERROR', 'format': 'red_format'})
worksheet5.conditional_format(columns_dataframe, {'type': 'cell', 'criteria': 'equal to', 'value': 'ERROR', 'format': 'red_format'})


workbook.close()


