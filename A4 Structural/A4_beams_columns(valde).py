###############################################
#### Initial code info                     ####
###############################################

# To run code following 3 packages must be installed to pip (pip install [package]): 
# (1) pandas   (2) xlwt   (3) openpyxl



###############################################
#### Model load                            ####
###############################################

import tkinter as tk
from tkinter.filedialog import askopenfilename
tk.Tk().withdraw() # part of the import if you are not using other tkinter functions

Ifc_path = askopenfilename() #The path the the Ifc file is saved at Ifc_path

import ifcopenshell
file = ifcopenshell.open(Ifc_path)  #The selected Ifc file gets imported.


###############################################
#### Element list Extraction               ####
###############################################

# Creating element vectors for later property extraction.
beams = file.by_type('IfcBeam')
columns = file.by_type('IfcColumn') 
walls = file.by_type('IfcWallStandardCase') #Often outher walls
walls.append(file.by_type('IfcWall')) # Often inner walls
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
    # beam_El_name,  beam_Material,  beam_Length,  beam_placement,  beam_level,  beam_crossection

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
                for property in property_set.HasProperties:
                    if property.Name == "Reference Level":
                        level_beam = str(property.NominalValue.wrappedValue)
                        beam_level.append(level_beam)

    #Extracting Length
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "Length":
                        length_beam = str(round(property.NominalValue.wrappedValue, 2))
                        beam_Length.append(length_beam)

    #Extracting Material
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Materials and Finishes":
                for property in property_set.HasProperties:
                    if property.Name == "Beam Material":
                        material_beam = str(property.NominalValue.wrappedValue)
                        beam_Material.append(material_beam)


    #Extracting Crossection bf
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "bf":
                        crossection_beam_bf = str(property.NominalValue.wrappedValue)
                        beam_crossection_bf.append(crossection_beam_bf)


    #Extracting Crossection d
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "d":
                        crossection_beam_d = str(property.NominalValue.wrappedValue)
                        beam_crossection_d.append(crossection_beam_d)

    #Extracting Crossection k
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "k":
                        crossection_beam_k = str(property.NominalValue.wrappedValue)
                        beam_crossection_k.append(crossection_beam_k)

    #Extracting Crossection kr
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "kr":
                        crossection_beam_kr = str(property.NominalValue.wrappedValue)
                        beam_crossection_kr.append(crossection_beam_kr)

    #Extracting Crossection tf
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "tf":
                        crossection_beam_tf = str(property.NominalValue.wrappedValue)
                        beam_crossection_tf.append(crossection_beam_tf)

    #Extracting Crossection tw
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "tw":
                        crossection_beam_tw = str(round(property.NominalValue.wrappedValue, 2))
                        beam_crossection_tw.append(crossection_beam_tw)


    #Extracting Object coordinates
    xval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[0],3)
    beam_placement_x.append(xval)
    yval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[1],3)
    beam_placement_y.append(yval)
    zval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[2],3)
    beam_placement_z.append(zval)


print(beam_num)
print(beam_El_name)
print(beam_Material)
print(beam_Length)
print(beam_placement_x)
print(beam_placement_y)
print(beam_placement_z)
print(beam_level)
print(beam_crossection_bf)
print(beam_crossection_d)
print(beam_crossection_k)
print(beam_crossection_kr)
print(beam_crossection_tf)
print(beam_crossection_tw)



######## COLUMNS:
# NOTE: 
  # Expected information vectors named with associated entity:
    # column_El_name,  column_Material,  column_Length,  column_placement,  column_level,  column_crossection
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
column_crossection_r = []
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
            if property_set.Name == "PSet_BuildingElementCommon":
                for property in property_set.HasProperties:
                    if property.Name == "Level":
                        level_column = str(property.NominalValue.wrappedValue)
                        column_level.append(level_column)

    #Extracting Length
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "Length":
                        length_column = str(round(property.NominalValue.wrappedValue, 2))
                        column_Length.append(length_column)

    #Extracting Material
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Materials and Finishes":
                for property in property_set.HasProperties:
                    if property.Name == "Beam Material":
                        material_beam = str(property.NominalValue.wrappedValue)
                        beam_Material.append(material_beam)


    #Extracting Crossection bf
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "bf":
                        crossection_beam_bf = str(property.NominalValue.wrappedValue)
                        beam_crossection_bf.append(crossection_beam_bf)


    #Extracting Crossection d
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "d":
                        crossection_beam_d = str(property.NominalValue.wrappedValue)
                        beam_crossection_d.append(crossection_beam_d)

    #Extracting Crossection k
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "k":
                        crossection_beam_k = str(property.NominalValue.wrappedValue)
                        beam_crossection_k.append(crossection_beam_k)

    #Extracting Crossection kr
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "kr":
                        crossection_beam_kr = str(property.NominalValue.wrappedValue)
                        beam_crossection_kr.append(crossection_beam_kr)

    #Extracting Crossection tf
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "tf":
                        crossection_beam_tf = str(property.NominalValue.wrappedValue)
                        beam_crossection_tf.append(crossection_beam_tf)

    #Extracting Crossection tw
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "tw":
                        crossection_beam_tw = str(property.NominalValue.wrappedValue)
                        beam_crossection_tw.append(crossection_beam_tw)


    #Extracting Object coordinates
    xval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[0],3)
    beam_placement_x.append(xval)
    yval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[1],3)
    beam_placement_y.append(yval)
    zval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[2],3)
    beam_placement_z.append(zval)

print(column_num)
print(column_El_name)
print(column_Material)
print(column_Length)
print(column_placement_x)
print(column_placement_y)
print(column_placement_z)
print(column_level)
print(column_crossection_b)
print(column_crossection_h)
print(column_crossection_r)
print(column_crossection_tf)
print(column_crossection_tw)





######## WALLS:
# NOTE: 
  # Expected information vectors named with associated entity:
    # 





######## SLABS:



###############################################
#### Converting Panda to Excel             ####
###############################################

# import pandas as pd
# import openpyxl   (is only used to add data to an existing Excel file)


# Data Frame is working on a matrix basis, Index = row titles, Columns = column titles
# every [] is a column in the matrix
# This step must be automated so it follows the different structural elements. 

#### EXAMPLE:
# df = pd.DataFrame([[11, 21, 31], [12, 22, 32], [31, 32, 33]],
#                    index=['one', 'two', 'three'], columns=['a', 'b', 'c'])


# creating Data Frame for the beams
# df_beams = pd.DataFrame(
#     [[beam_El_name], [beam_Material], [beam_Length], [beam_placement], [beam_crossection]], 
#     index=[beam_num] , 
#     colums=['Element name', 'Material', 'Length', 'Placement', 'crossection']
# )

# creating Data Frame for the columns
# df_columns = pd.DataFrame(
#     [[column_El_name], [column_Material], [column_Length], [column_placement], [column_crossection]], 
#     index=[column_num] , 
#     colums=['Element name', 'Material', 'Length', 'Placement', 'crossection']
# )

# creating Data Frame for the walls
# df_walls = pd.DataFrame(
#     [[wall_El_name], [wall_Material], [wall_Length], [wall_with], [wall_placement], [wall_Voids], [wall_Voidarea] ], 
#     index=[wall_num] , 
#     colums=['Element name', 'Material', 'Length', 'with', 'Placement']
# )

# creating Data Frame for the slaps
# df_slabs = pd.DataFrame(
#     [[slab_El_name], [slab_Material], [slab_Lenght], [slab_With], [slab_Placement]], 
#     index=[slab_num] , 
#     colums=['Element name', 'Material', 'Length', 'With', 'Placement']
# )

# Creating Excel file consisting the Data Frames

# with pd.ExcelWriter('pandas_to_excel.xlsx') as writer:
#     df_beams.to_excel(writer, sheet_name='Beams', index=True, header=True)
#     df_columns.to_excel(writer, sheet_name='Columns', index=True, header=True)
#     df_walls.to_excel(writer, sheet_name='Walls', index=True, header=True)
#     df_slabs.to_excel(writer, sheet_name='Slabs', index=True, header=True)