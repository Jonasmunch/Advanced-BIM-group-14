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

#######################
####### BEAMS: ########
#######################

num = 0

beam_num = []
beam_El_name = []
beam_Type = []
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

    # Beam numbering:
    num = num+1
    b_num = "beam"+str(num)
    beam_num.append(b_num)

    # Element name:
    name_b = beam.Name
    beam_El_name.append("ERROR")  if  name_b == ""  else beam_El_name.append(name_b)

    #Extracting element properties:
    level_b = [] ; length_b = [] ; type_b = [] ; bf = [] ; d = [] ; k = [] ; kr = [] ; tf = [] ; tw = []
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            
            #Extracting Reference level
            if property_set.Name == "PSet_Revit_Constraints":
                for property in property_set.HasProperties:
                    if property.Name == "Reference Level":
                        level_b = property.NominalValue.wrappedValue
                
            #Extracting Length
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "Length":
                        length_b = round(property.NominalValue.wrappedValue, 2)              

            #Extracting Type
            if property_set.Name == "PSet_Revit_Materials and Finishes":
                for property in property_set.HasProperties:
                    if property.Name == "Beam Material":
                        type_b = property.NominalValue.wrappedValue

            #Extracting Crossection
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    # bf
                    if property.Name == "bf":
                        bf = round(property.NominalValue.wrappedValue, 2)
                    # d
                    elif property.Name == "d":
                        d = round(property.NominalValue.wrappedValue, 2)
                    # k
                    elif property.Name == "k":
                        k = round(property.NominalValue.wrappedValue, 2)
                    # kr
                    elif property.Name == "kr":
                        kr = round(property.NominalValue.wrappedValue, 2)
                    # tf
                    elif property.Name == "tf":
                        tf = round(property.NominalValue.wrappedValue, 2)
                    # tw
                    elif property.Name == "tw":
                        tw = round(property.NominalValue.wrappedValue, 2)

    # Evaluating for missing properties:
    beam_level.append("ERROR")  if  level_b == ""   else  beam_level.append(level_b)
    beam_Length.append("ERROR") if  length_b == ""  else  beam_Length.append(length_b)
    beam_Type.append("ERROR")   if  type_b == ""    else  beam_Type.append(type_b)
    beam_crossection_bf.append("ERROR")  if  bf == ""  else beam_crossection_bf.append(bf)
    beam_crossection_d.append("ERROR")   if  d == ""   else beam_crossection_d.append(d)
    beam_crossection_k.append("ERROR")   if  k == ""   else beam_crossection_k.append(k)
    beam_crossection_kr.append("ERROR")  if  kr == ""  else beam_crossection_kr.append(kr)
    beam_crossection_tf.append("ERROR")  if  tf == ""  else beam_crossection_tf.append(tf)
    beam_crossection_tw.append("ERROR")  if  tw == ""  else beam_crossection_tw.append(tw)


    #Extracting Object coordinates
    xval = [] ; yval = [] ; zval = []
    xval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[0],3)
    yval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[1],3)
    zval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[2],3)

    # Evaluating for missing properties:
    beam_placement_x.append("ERROR")  if  xval == ""  else beam_placement_x.append(xval)
    beam_placement_y.append("ERROR")  if  yval == ""  else beam_placement_y.append(yval)
    beam_placement_z.append("ERROR")  if  zval == ""  else beam_placement_z.append(zval)

dataframe_beams = [beam_num, beam_El_name, beam_level, beam_Type, beam_Length, beam_crossection_bf, beam_crossection_d, beam_crossection_k, beam_crossection_kr, beam_crossection_tf, beam_crossection_tw, beam_placement_x, beam_placement_y, beam_placement_z]
beams_dataframe = numpy.transpose(dataframe_beams)



##########################
######## COLUMNS: ########
##########################

num = 0

column_num = []
column_El_name = []
column_Type = []
column_Length = []
column_placement_x = []
column_placement_y = []
column_placement_z = []
column_level = []
column_crossection_bf = []
column_crossection_d = []
column_crossection_k = []
column_crossection_kr = []
column_crossection_tf = []
column_crossection_tw = []

for column in columns:

    # Column numbering:
    num = num+1
    c_num = "column"+str(num)
    column_num.append(c_num)

    # Element name:
    name_c = column.Name
    column_El_name.append("ERROR")  if  name_c == ""  else column_El_name.append(name_c)

    #Extracting element properties:
    level_c = [] ; length_c = [] ; type_c = [] ; bf = [] ; d = [] ; k = [] ; kr = [] ; tf = [] ; tw = []
    #
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            
            #Extracting Reference level
            if property_set.Name == "PSet_Revit_Constraints":
                for property in property_set.HasProperties:
                    if property.Name == "Reference Level":
                        level_c = property.NominalValue.wrappedValue
                
            #Extracting Length
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "Length":
                        length_c = round(property.NominalValue.wrappedValue, 2)              

            #Extracting Type
            if property_set.Name == "PSet_Revit_Materials and Finishes":
                for property in property_set.HasProperties:
                    if property.Name == "Beam Material":
                        type_c = property.NominalValue.wrappedValue

            #Extracting Crossection
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    # bf
                    if property.Name == "bf":
                        bf = round(property.NominalValue.wrappedValue, 2)
                    # d
                    elif property.Name == "d":
                        d = round(property.NominalValue.wrappedValue, 2)
                    # k
                    elif property.Name == "k":
                        k = round(property.NominalValue.wrappedValue, 2)
                    # kr
                    elif property.Name == "kr":
                        kr = round(property.NominalValue.wrappedValue, 2)
                    # tf
                    elif property.Name == "tf":
                        tf = round(property.NominalValue.wrappedValue, 2)
                    # tw
                    elif property.Name == "tw":
                        tw = round(property.NominalValue.wrappedValue, 2)

    # Evaluating for missing properties:
    column_level.append("ERROR")  if  level_c == ""   else  column_level.append(level_c)
    column_Length.append("ERROR") if  length_c == ""  else  column_Length.append(length_c)
    column_Type.append("ERROR")   if  type_c == ""    else  column_Type.append(type_c)
    column_crossection_bf.append("ERROR")  if  bf == ""  else column_crossection_bf.append(bf)
    column_crossection_d.append("ERROR")   if  d == ""   else column_crossection_d.append(d)
    column_crossection_k.append("ERROR")   if  k == ""   else column_crossection_k.append(k)
    column_crossection_kr.append("ERROR")  if  kr == ""  else column_crossection_kr.append(kr)
    column_crossection_tf.append("ERROR")  if  tf == ""  else column_crossection_tf.append(tf)
    column_crossection_tw.append("ERROR")  if  tw == ""  else column_crossection_tw.append(tw)


    #Extracting Object coordinates
    xval = [] ; yval = [] ; zval = []
    xval = round(column.ObjectPlacement.RelativePlacement.Location.Coordinates[0],3)
    yval = round(column.ObjectPlacement.RelativePlacement.Location.Coordinates[1],3)
    zval = round(column.ObjectPlacement.RelativePlacement.Location.Coordinates[2],3)

    # Evaluating for missing properties:
    column_placement_x.append("ERROR")  if  xval == ""  else column_placement_x.append(xval)
    column_placement_y.append("ERROR")  if  yval == ""  else column_placement_y.append(yval)
    column_placement_z.append("ERROR")  if  zval == ""  else column_placement_z.append(zval)

dataframe_columns = [column_num, column_El_name, column_level, column_Type, column_Length, column_crossection_bf, column_crossection_d, column_crossection_k, column_crossection_kr, column_crossection_tf, column_crossection_tw, column_placement_x, column_placement_y, column_placement_z]
columns_dataframe = numpy.transpose(dataframe_columns)


########################
######## WALLS: ########
########################

# Creating empty property vectors:
wall_num = [] 
wall_El_name = [] 
wall_Level = [] 
wall_Height =[] 
wall_Lenght = [] 
wall_Width =[] 
wall_Volume =[] 
wall_Type = [] 
wall_xval = [] 
wall_yval = [] 
wall_zval = []

num = 0

for wall in Walls:
#
    # Wall numbering:
    num = num+1
    w_num = "wall"+str(num)
    wall_num.append(w_num)

    # Element name:
    name_w = wall.Name
    wall_El_name.append("ERROR")  if  name_w == ""  else  wall_El_name.append(name_w)
# 
#
    height_w =[] ; length_w = [] ; thickness_w = [] ; volume_w = [] ; level_w = []

    for definition in wall.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            #
            #Extracting Height of walls
            if property_set.Name == "PSet_Revit_Constraints":
                for property in property_set.HasProperties:
                    if property.Name == "Unconnected Height":
                        height_w = (round(property.NominalValue.wrappedValue, 2))
            #
            #Extracting length of wall
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "Length":
                        length_w = (round(property.NominalValue.wrappedValue, 2))
            #
            #Extracting thinkness of wall
            if property_set.Name == "PSet_Revit_Type_Construction":
                for property in property_set.HasProperties:
                    if property.Name == "Width":
                        thickness_w = (round(property.NominalValue.wrappedValue, 2))
            #
            #Extracting volume of wall
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "Volume":
                        volume_w = (round(property.NominalValue.wrappedValue, 2))
            #
            #Extracting reference level of wall
            if property_set.Name == "PSet_Revit_Constraints":
                for property in property_set.HasProperties:
                    if property.Name == "Base Constraint":
                        level_w = property.NominalValue.wrappedValue
    #
        # Evaluating for missing properties:
    wall_Height.append("ERROR")  if height_w == ""   else  wall_Height.append(height_w)
    wall_Lenght.append("ERROR")  if length_w == ""   else  wall_Lenght.append(length_w)
    wall_Width.append("ERROR")   if thickness_w == ""   else wall_Width.append(thickness_w)
    wall_Volume.append("ERROR")  if volume_w == ""   else  wall_Volume.append(volume_w)
    wall_Level.append("ERROR")   if level_w == ""    else  wall_Level.append(level_w)

    #Extracting wall type
    for RAM in wall.HasAssociations:
        type_w = RAM.RelatingMaterial.ForLayerSet.LayerSetName
    # Evaluating for missing properties:
    wall_Type.append("ERROR")  if type_w == ""  else wall_Type.append(type_w)

    # Extracting coordinates
    xval = [] ; yval = [] ; zval = []
    xval = round(wall.ObjectPlacement.RelativePlacement.Location.Coordinates[0],3)
    yval = round(wall.ObjectPlacement.RelativePlacement.Location.Coordinates[1],3)
    zval = round(wall.ObjectPlacement.RelativePlacement.Location.Coordinates[2],3)
    
    # Evaluating for missing properties in coordinates:
    wall_xval.append("ERROR")  if xval == ""  else  wall_xval.append(xval)
    wall_yval.append("ERROR")  if yval == ""  else  wall_yval.append(yval)
    wall_zval.append("ERROR")  if zval == ""  else  wall_zval.append(zval)


# Creating collective dataframe matrix:
dataframe_walls = [wall_num, wall_El_name, wall_Level, wall_Type, wall_Width, wall_Height, wall_Lenght, wall_Volume, wall_xval, wall_yval, wall_zval]
walls_dataframe = numpy.transpose(dataframe_walls)



########################
######## SLABS: ########
########################

slabs_structural = [] 
slab_num = [] 
slab_El_name = [] 
slab_Perimeter = [] 
slab_Type = []
slab_Area = [] 
slab_Volume = [] 
slab_Thickness = [] 
slab_Level = []
slab_xval = [] 
slab_yval = [] 
slab_zval = []
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

    # slab numbering:
    num = num+1
    slab_num.append("Slab"+str(num))

    # Element name:
    name_s = slab.Name
    slab_El_name.append("ERROR")  if  name_s == ""  else  slab_El_name.append(name_s)
    
    # Extracting dimensional properties:
    perimeter_s = [] ; area_s = [] ; volume_s = [] ; thickness_s = [] ; level_s = []
    
    for definition in slab.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    # Perimeter properties:
                    if property.Name == "Perimeter":
                        perimeter_s = round(property.NominalValue.wrappedValue, 2)
                    # Area Properties:
                    elif property.Name == "Area":
                        area_s = round(property.NominalValue.wrappedValue, 2)
                    # Volume Properties:
                    elif property.Name == "Volume":
                        volume_s = round(property.NominalValue.wrappedValue, 2)
                    # Thickness properties:
                    elif property.Name == "Thickness":
                        thickness_s = round(property.NominalValue.wrappedValue, 2)
            
            # Extracting reference level:
            if property_set.Name == "PSet_Revit_Constraints":
                for property in property_set.HasProperties:
                    if property.Name == "Level":
                        level_s = property.NominalValue.wrappedValue
                
    # Evaluating for missing properties:
    slab_Perimeter.append("ERROR")  if  perimeter_s == ""  else  slab_Perimeter.append(perimeter_s)
    slab_Area.append("ERROR")       if  area_s == ""       else  slab_Area.append(area_s)
    slab_Volume.append("ERROR")     if  volume_s == ""     else  slab_Volume.append(volume_s)
    slab_Thickness.append("ERROR")  if  thickness_s == ""  else  slab_Thickness.append(thickness_s)
    slab_Level.append("ERROR")      if  level_s == ""      else  slab_Level.append(level_s)
    
    # Extracting Material:
    type_s = []
    for RAM in slab.HasAssociations:
        type_s = RAM.RelatingMaterial.ForLayerSet.LayerSetName
    # Evaluating for missing properties:
    slab_Type.append("ERROR")  if type_s == ""  else slab_Type.append(type_s)
    
    # Extracting coordinates
    xval = [] ; yval = [] ; zval = []
    xval = round(slab.ObjectPlacement.RelativePlacement.Location.Coordinates[0],3)
    yval = round(slab.ObjectPlacement.RelativePlacement.Location.Coordinates[1],3)
    zval = round(slab.ObjectPlacement.RelativePlacement.Location.Coordinates[2],3)
    
    # Evaluating for missing properties in coordinates:
    slab_xval.append("ERROR")  if xval == ""  else  slab_xval.append(xval)
    slab_yval.append("ERROR")  if yval == ""  else  slab_yval.append(yval)
    slab_zval.append("ERROR")  if zval == ""  else  slab_zval.append(zval)


dataframe_slabs = [slab_num, slab_El_name, slab_Level, slab_Type, slab_Thickness, slab_Area, slab_Perimeter, slab_Volume, slab_xval, slab_yval, slab_zval]
slabs_dataframe = numpy.transpose(dataframe_slabs)



###############################################
#### Naming output file                    ####
###############################################

from tkinter import simpledialog

ROOT = tk.Tk()
ROOT.withdraw()
# Opening the nameing dialog
Output_FileN = simpledialog.askstring(title="Test", prompt="Filename of outputfile")  
Output_Filename = str(Output_FileN)+'.xlsx'


###############################################
#### Creating Excel document               ####
###############################################


###########
# Creating workbook and worksheets:
workbook = xlsxwriter.Workbook(Output_Filename)
worksheet1 = workbook.add_worksheet('Info')
worksheet2 = workbook.add_worksheet('Slabs')
worksheet3 = workbook.add_worksheet('Walls')
worksheet4 = workbook.add_worksheet('Beams')
worksheet5 = workbook.add_worksheet('Columns')


######
# Formatting:
red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
green_format = workbook.add_format({'bg_color':   '#C6EFCE', 'font_color': '#006100'})


######
# Creating tables within the sheets:

# SLABS:
slabs_rows = len(slabs_structural)
slabs_columns = len(dataframe_slabs)-1

if slabs_rows <= 0:
    slabs_rows = 1
    slabs_columns = 13
    worksheet2.add_table(0,0, 1, 10, { 
        'data': [],
        'style': 'Table Style Light 11',
        'name': "Slabs_data",
        'columns': [{'header': 'Element'}, {'header': 'Element Name'}, {'header': 'Ref_level'}, {'header': 'Type'}, 
        {'header': 'Thickness [m]'}, {'header': 'Area [m^2]'}, {'header': 'Perimeter [m]'}, {'header': 'Volume [m^3]'}, {'header': 'Placement_x-val'},
        {'header': 'Placement_y-val'}, {'header': 'Placement_z-val'},]
    })  
else:
    worksheet2.add_table(0,0, slabs_rows, slabs_columns, { 
    'data': slabs_dataframe,
    'style': 'Table Style Light 11',
    'name': "Slabs_data",
    'columns': [{'header': 'Element'}, {'header': 'Element Name'}, {'header': 'Ref_level'}, {'header': 'Type'}, 
    {'header': 'Thickness [m]'}, {'header': 'Area [m^2]'}, {'header': 'Perimeter [m]'}, {'header': 'Volume [m^3]'}, {'header': 'Placement_x-val'},
    {'header': 'Placement_y-val'}, {'header': 'Placement_z-val'},]
}) 

# WALLS:
walls_rows = len(Walls)
walls_columns = len(dataframe_walls)-1

if walls_rows <= 0:
    walls_rows = 1
    walls_columns = 13
    worksheet3.add_table(0,0,1,10, { 
        'data': [],
        'style': 'Table Style Light 11',
        'name': "Walls_data",
        'columns': [{'header': 'Element'}, {'header': 'Element Name'}, {'header': 'Ref_level'}, {'header': 'Material'}, 
        {'header': 'Thickness'}, {'header': 'Area'}, {'header': 'Perimeter'}, {'header': 'Volume'}, {'header': 'Placement_x-val'},
        {'header': 'Placement_y-val'}, {'header': 'Placement_z-val'},]
    })
else:
    worksheet3.add_table(0,0,walls_rows,walls_columns, { 
        'data': walls_dataframe,
        'style': 'Table Style Light 11',
        'name': "Walls_data",
        'columns': [{'header': 'Element'}, {'header': 'Element Name'}, {'header': 'Ref_level'}, {'header': 'Material'}, 
        {'header': 'Thickness'}, {'header': 'Area'}, {'header': 'Perimeter'}, {'header': 'Volume'}, {'header': 'Placement_x-val'},
        {'header': 'Placement_y-val'}, {'header': 'Placement_z-val'},]
    })

# BEAMS:
beams_rows = len(beams)
beams_columns = len(dataframe_beams)-1

if beams_rows <= 0:
    beams_rows = 1
    beams_columns = 13
    worksheet4.add_table(0,0,1,13, { 
        'data': [],
        'style': 'Table Style Light 11',
        'name': "Beams_data",
        'columns': [{'header': 'Element'}, {'header': 'Element Name'}, {'header': 'Ref_level'}, {'header': 'Material'}, 
        {'header': 'Length'}, {'header': 'Width of flange, bf'}, {'header': 'Height of web, d'}, {'header': 'k'}, {'header': 'kr'},
        {'header': 'Thickness of flange, tf'}, {'header': 'Thickness of web, tw'}, {'header': 'Placement_x-val'}, {'header': 'Placement_y-val'}, 
        {'header': 'Placement_z-val'},]
    })
else:
    worksheet4.add_table(0,0,beams_rows,beams_columns, { 
        'data': beams_dataframe,
        'style': 'Table Style Light 11',
        'name': "Beams_data",
        'columns': [{'header': 'Element'}, {'header': 'Element Name'}, {'header': 'Ref_level'}, {'header': 'Material'}, 
        {'header': 'Length'}, {'header': 'Width of flange, bf'}, {'header': 'Height of web, d'}, {'header': 'k'}, {'header': 'kr'},
        {'header': 'Thickness of flange, tf'}, {'header': 'Thickness of web, tw'}, {'header': 'Placement_x-val'}, {'header': 'Placement_y-val'}, 
        {'header': 'Placement_z-val'},]
    })

# COLUMNS:
columns_rows = len(columns)
columns_columns = len(dataframe_columns)-1

if columns_rows <= 0:
    columns_rows = 1
    columns_columns = 13
    worksheet5.add_table(0,0,columns_rows,columns_columns, { 
    'data': [["ERROR", "ERROR"]],
    'style': 'Table Style Light 11',
    'name': "Columns_data",
    'columns': [{'header': 'Element'}, {'header': 'Element Name'}, {'header': 'Ref_level'}, {'header': 'Material'}, 
    {'header': 'Length'}, {'header': 'Width, b'}, {'header': 'Height, h'}, {'header': 'k'}, {'header': 'kr'},
    {'header': 'Thickness of flange, tf'}, {'header': 'Thickness of web, tw'}, {'header': 'Placement_x-val'}, {'header': 'Placement_y-val'}, 
    {'header': 'Placement_z-val'},]
    })
else:
    worksheet5.add_table(0,0,columns_rows,columns_columns, { 
        'data': columns_dataframe,
        'style': 'Table Style Light 11',
        'name': "Columns_data",
        'columns': [{'header': 'Element'}, {'header': 'Element Name'}, {'header': 'Ref_level'}, {'header': 'Material'}, 
        {'header': 'Length'}, {'header': 'Width, b'}, {'header': 'Height, h'}, {'header': 'k'}, {'header': 'kr'},
        {'header': 'Thickness of flange, tf'}, {'header': 'Thickness of web, tw'}, {'header': 'Placement_x-val'}, {'header': 'Placement_y-val'}, 
        {'header': 'Placement_z-val'},]
    })

    
#####
# Conditional formatting: 

worksheet2.conditional_format(1,0,slabs_rows,slabs_columns,  {'type': 'text', 'criteria': 'containing', 'value': 'ERROR', 'format': red_format})
worksheet2.conditional_format(1,0,slabs_rows,slabs_columns, {'type': 'text', 'criteria': 'not containing', 'value': 'ERROR', 'format': green_format})

worksheet3.conditional_format(1,0,walls_rows,walls_columns,  {'type': 'text', 'criteria': 'containing', 'value': 'ERROR', 'format': red_format})
worksheet3.conditional_format(1,0,walls_rows,walls_columns, {'type': 'text', 'criteria': 'not containing', 'value': 'ERROR', 'format': green_format})

worksheet4.conditional_format(1,0,beams_rows,beams_columns,  {'type': 'text', 'criteria': 'containing', 'value': 'ERROR', 'format': red_format})
worksheet4.conditional_format(1,0,beams_rows,beams_columns, {'type': 'text', 'criteria': 'not containing', 'value': 'ERROR', 'format': green_format})

worksheet5.conditional_format(1,0,columns_rows,columns_columns,  {'type': 'text', 'criteria': 'containing', 'value': 'ERROR', 'format': red_format})
worksheet5.conditional_format(1,0,columns_rows,columns_columns, {'type': 'text', 'criteria': 'not containing', 'value': 'ERROR', 'format': green_format})


#####
# Creating content for Info-sheet:

#
data = [ ["Category" , "Slabs" , "Walls" , "Beams" , "Columns" , "Total" ],
["Elements", '=COUNTA(Slabs_data[Element])' , '=COUNTA(Walls_data[Element])' , '=COUNTA(Beams_data[Element])' , '=COUNTA(Columns_data[Element])' , '=SUM(L4:L7)' ],
["Enteties", '=COUNTA(Slabs_data[])' , '=COUNTA(Walls_data[])' , '=COUNTA(Beams_data[])' , '=COUNTA(Columns_data[])' , '=SUM(M4:M7)-N8' ],
["Errors", '=COUNTIF(Slabs_data[], "ERROR")' , '=COUNTIF(Walls_data[], "ERROR")' , '=COUNTIF(Beams_data[], "ERROR")' , '=COUNTIF(Columns_data[], "ERROR")' , '=SUM(N4:N7)'],
]

worksheet1.write_column('K3', data[0])
worksheet1.write_column('L3', data[1])
worksheet1.write_column('M3', data[2])
worksheet1.write_column('N3', data[3])


# Preparing chards:
pie_chart = workbook.add_chart({'type': 'pie'})
column_chart = workbook.add_chart({'type': 'column'})

# Creating pie chart:
pie_chart.add_series({
    'name': "Error destribution",
    'categories': '=Info!$M$3:$N$3',
    'values':     '=Info!$M$8:$N$8',
    'data_labels': {'percentage': True},
    'points': [
        {'fill': {'color': 'green'}},
        {'fill': {'color': 'red'}},
    ],
})

# Creating column chart:
column_chart.add_series({
    'name': "Correct",
    'categories': '=Info!$K$4:$K$7',
    'values': '=Info!$M$4:$M$7',
    'fill': {'color': 'green'},
})
column_chart.add_series({
    'name': "Error",
    'values': '=Info!$N$4:$N$7',
    'fill': {'color': 'red'},
})
column_chart.set_title({
    'name': 'Category results',
    'name_font': {
        'name': 'Calibri',
        'color': 'black',
    },
})

# Insetting charts into Info sheet:
worksheet1.insert_chart('B3', pie_chart)
worksheet1.insert_chart('B18', column_chart)

workbook.close()


