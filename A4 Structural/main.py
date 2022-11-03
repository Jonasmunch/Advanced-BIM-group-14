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
    # beam_El_name,  beam_Material,  beam_Lenght,  beam_placement,  beam_crossection



######## COLUMNS:
# NOTE: 
  # Expected information vectors named with associated entity:
    # column_El_name,  column_Material,  column_Lenght,  column_placement,  column_crossection




######## WALLS:
# NOTE: 
  # Expected information vectors named with associated entity:
    # 





######## SLABS:



###############################################
#### Converting Panda to Excel             ####
###############################################

import pandas as pd
# import openpyxl   (is only used to add data to an existing Excel file)


# Data Frame is working on a matrix basis, Index = row titles, Columns = column titles
# every [] is a column in the matrix
# This step must be automated so it follows the different structural elements. 

#### EXAMPLE:
# df = pd.DataFrame([[11, 21, 31], [12, 22, 32], [31, 32, 33]],
#                    index=['one', 'two', 'three'], columns=['a', 'b', 'c'])


# creating Data Frame for the beams
df_beams = pd.DataFrame(
    [[beam_El_name], [beam_Material], [beam_Lenght], [beam_placement], [beam_crossection]], 
    index=[beam_num] , 
    colums=['Element name', 'Material', 'Lenght', 'Placement', 'crossection']
)

# creating Data Frame for the columns
df_columns = pd.DataFrame(
    [[column_El_name], [column_Material], [column_Lenght], [column_placement], [column_crossection]], 
    index=[column_num] , 
    colums=['Element name', 'Material', 'Lenght', 'Placement', 'crossection']
)

# creating Data Frame for the walls
df_walls = pd.DataFrame(
    [[wall_El_name], [wall_Material], [wall_Lenght], [wall_with], [wall_placement], [wall_Voids], [wall_Voidarea] ], 
    index=[wall_num] , 
    colums=['Element name', 'Material', 'Lenght', 'with', 'Placement']
)

# creating Data Frame for the slaps
df_slabs = pd.DataFrame(
    [[slab_El_name], [slab_Material], [slab_Lenght], [slab_With], [slab_Placement]], 
    index=[slab_num] , 
    colums=['Element name', 'Material', 'Lenght', 'With', 'Placement']
)

# Creating Excel file consisting the Data Frames

with pd.ExcelWriter('pandas_to_excel.xlsx') as writer:
    df_beams.to_excel(writer, sheet_name='Beams', index=True, header=True)
    df_columns.to_excel(writer, sheet_name='Columns', index=True, header=True)
    df_walls.to_excel(writer, sheet_name='Walls', index=True, header=True)
    df_slabs.to_excel(writer, sheet_name='Slabs', index=True, header=True)


