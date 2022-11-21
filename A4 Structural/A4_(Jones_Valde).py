''' Editted by Joakim Mørk, Valdemar Rasmussen, Jonas Munch, Oscar Hansen 2022 '''

import math

import tkinter as tk
from tkinter.filedialog import askopenfilename
tk.Tk().withdraw() # part of the import if you are not using other tkinter functions

Ifc_path = askopenfilename() #The path the the Ifc file is saved at Ifc_path

import ifcopenshell
file = ifcopenshell.open(Ifc_path)  #The selected Ifc file gets imported.



#########################
#   Jones og Valdes     #
#########################


# Creating element vectors for later property extraction.
beams = file.by_type('IfcBeam')
if len(beams) == 0:
    global_call = "No data have been found"

columns = file.by_type('IfcColumn') 
if len(columns) == 0:
    global_call = "No data have been found"

Walls = file.by_type('IfcWall')
if len(Walls) == 0:
    global_call = "No data have been found"

# Creating empty vectors, for easy element numbering.
beam_num = []
column_num = []
wall_num = []


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
            call_a = "ERROR"
            if property_set.Name == "PSet_Revit_Constraints":
                for property in property_set.HasProperties:
                    if property.Name == "Reference Level":
                        level_beam = str(property.NominalValue.wrappedValue)
                        beam_level.append(level_beam)
                    else:
                        beam_level.append(call_a)
            else:
                beam_level.append(call_a)

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


# print(beam_num)
# print(beam_El_name)
# print(beam_Material)
# print(beam_Length)
# print(beam_placement_x)
# print(beam_placement_y)
# print(beam_placement_z)
# print(beam_level)
# print(beam_crossection_bf)
# print(beam_crossection_d)
# print(beam_crossection_k)
# print(beam_crossection_kr)
# print(beam_crossection_tf)
# print(beam_crossection_tw)



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
            call_a = "ERROR"
            if property_set.Name == "PSet_BuildingElementCommon":
                for property in property_set.HasProperties:
                    if property.Name == "Level":
                        level_column = str(property.NominalValue.wrappedValue)
                        column_level.append(level_column)
                    else:
                        column_level.append(call_a)
            else:
                column_level.append(call_a)

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
                    if property.Name == "Column Material":
                        material_column = str(property.NominalValue.wrappedValue)
                        column_Material.append(material_column)


    #Extracting Crossection b
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "b":
                        crossection_column_b = str(property.NominalValue.wrappedValue)
                        column_crossection_b.append(crossection_column_b)


    #Extracting Crossection h
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "h":
                        crossection_column_h = str(property.NominalValue.wrappedValue)
                        column_crossection_h.append(crossection_column_h)

    #Extracting Crossection k
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "k":
                        crossection_column_k = str(property.NominalValue.wrappedValue)
                        column_crossection_k.append(crossection_column_k)

    #Extracting Crossection kr
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "kr":
                        crossection_column_kr = str(property.NominalValue.wrappedValue)
                        column_crossection_kr.append(crossection_column_kr)

    #Extracting Crossection tf
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "tf":
                        crossection_column_tf = str(property.NominalValue.wrappedValue)
                        column_crossection_tf.append(crossection_column_tf)

    #Extracting Crossection tw
    for definition in column.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "tw":
                        crossection_column_tw = str(property.NominalValue.wrappedValue)
                        column_crossection_tw.append(crossection_column_tw)


    #Extracting Object coordinates
    xval = round(column.ObjectPlacement.RelativePlacement.Location.Coordinates[0],3)
    column_placement_x.append(xval)
    yval = round(column.ObjectPlacement.RelativePlacement.Location.Coordinates[1],3)
    column_placement_y.append(yval)
    zval = round(column.ObjectPlacement.RelativePlacement.Location.Coordinates[2],3)
    column_placement_z.append(zval)

# print(column_num)
# print(column_El_name)
# print(column_Material)
# print(column_Length)
# print(column_placement_x)
# print(column_placement_y)
# print(column_placement_z)
print(column_level)
# print(column_crossection_b)
# print(column_crossection_h)
# print(column_crossection_r)
# print(column_crossection_tf)
# print(column_crossection_tw)



#########################################################
##### Lavet af OHH <3 Walls #####
#########################################################


num = 0

wall_Height =[]
wall_Lenght = []
wall_Width = []
wall_Volume = []
wall_Type = []


for wall in Walls:
#
    num = num+1
    w_num = "wall"+str(num)
    wall_num.append(w_num)

    #Extracting Height of walls 
    for definition in wall.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Constraints":
                for property in property_set.HasProperties:
                    if property.Name == "Unconnected Height":
                        UH_w = (round(property.NominalValue.wrappedValue, 2))
                        wall_Height.append(UH_w)
                        
    #Extracting length of wall
    for definition in wall.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "Length":
                        length_w = (round(property.NominalValue.wrappedValue, 2))
                        wall_Lenght.append(length_w)

    #Extracting thinkness of wall
    for definition in wall.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Type_Construction":
                for property in property_set.HasProperties:
                    if property.Name == "Width":
                        thickness_w = (round(property.NominalValue.wrappedValue, 2))
                        wall_Width.append(thickness_w)

    #Extracting volume of wall
    for definition in wall.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
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
 

###############################################
#### Name output file                      ####
###############################################
import tkinter as tk
from tkinter import simpledialog

ROOT = tk.Tk()

ROOT.withdraw()

Output_FileN = simpledialog.askstring(title="Test",                   # make the nameing dialog
                                  prompt="Filename of outputfile")  

Output_Filename = str(Output_FileN)+'.xlsx'


###############################################
#### Exporting the excel                   ####
###############################################
import pandas as pd
import xlsxwriter

# Beam data
if global_call == "":
    beam_Data = {"ERROR" : global_call}
    df_beams = pd.DataFrame(beam_Data)
else:
    beam_Data = {'Beam name' : beam_El_name, 'Length' : beam_Length, 'Material' : beam_Material, 'Flange width' : crossection_beam_bf, 'Web height' : crossection_beam_d, 'k' : crossection_beam_k, 'kr' : crossection_beam_kr, 'tf' : crossection_beam_tf, 'tw' : crossection_beam_tw, 'x-coordinate' : beam_placement_x, 'y-coordinate' : beam_placement_y, 'z-coordinate' : beam_placement_z}
    df_beams = pd.DataFrame(beam_Data)

# Column data
if global_call == "":
    column_Data = {"ERROR" : global_call}
    df_columns = pd.DataFrame(column_Data)
else:
    column_Data = {'Column name' : column_El_name, 'Length' : column_Length, 'Material' : column_Material, 'Flange width' : crossection_column_b, 'Web height' : crossection_column_h, 'k' : crossection_column_k, 'kr' : crossection_column_kr, 'tf' : crossection_column_tf, 'tw' : crossection_column_tw, 'x-coordinate' : column_placement_x, 'y-coordinate' : column_placement_y, 'z-coordinate' : column_placement_z}
    df_columns = pd.DataFrame(column_Data)

# Wall data
if global_call == "":
    Wall_Data = {"ERROR" : global_call}
    df_walls = pd.DataFrame(Wall_Data)
else:
    wall_Data = {'Wall Type' : wall_Type, 'Lenght' : wall_Lenght, 'Width' : wall_Width, 'Height' : wall_Height, 'Volume' : wall_Volume }
    df_walls = pd.DataFrame(wall_Data)

# Creating Excel file consisting the Data Frames
with pd.ExcelWriter(Output_Filename) as writer:
    df_beams.to_excel(writer, sheet_name='Beams', index=True, header=True)

with pd.ExcelWriter(Output_Filename) as writer:
    df_columns.to_excel(writer, sheet_name='Columns', index=True, header=True)

with pd.ExcelWriter(Output_Filename) as writer:
    df_walls.to_excel(writer, sheet_name='Walls', index=True, header=True)
