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
# Creating empty property vectors:
slabs_structural = [] ; slab_num = [] ; slab_El_name = [] ; slab_Perimeter = [] ; slab_Material = []
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
            slab_Material.append(call_a)
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


dataframe = [slab_num, slab_El_name, slab_Level, slab_Material, slab_Thickness, slab_Area, slab_Perimeter, slab_Volume, slab_xval, slab_yval, slab_zval]
slab_dataframe = numpy.transpose(dataframe)


###############################################
#### Creating Excel document               ####
###############################################




