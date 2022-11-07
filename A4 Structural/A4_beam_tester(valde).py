''' written by Tim McGinley 2022 '''
''' Editted by Joakim Mørk, Valdemar Rasmussen, Jonas Munch, Oscar Hansen 2022 '''

# Additions made to HTMLBuild.py: 
# - Beam entity loader in 'writeCustomHTML' function, and a structural subfolder to withhold information on structural elements (114-126). 
# - 'ClassifyBeams' function to load in beam entities (192-237).

#import tkinter as tk
#from tkinter.filedialog import askopenfilename
#tk.Tk().withdraw() # part of the import if you are not using other tkinter functions

#Ifc_path = askopenfilename() #The path the the Ifc file is saved at Ifc_path

import ifcopenshell
file = ifcopenshell.open("C:/Users/valde/OneDrive - Danmarks Tekniske Universitet/Skole/Bachelor/7. semester/41934 Adv BIM/Duplex_A_20110907.ifc")  #The selected Ifc file gets imported.

##################################################
##### Searching for structural elements      #####
##################################################

Beams = file.by_type("IfcBeam")
Columns = file.by_type("IfcColumn")

#### testing for possible errors: ####

# if len(Columns) <= 0:    #logical testing if 'Columns' string is of lenght >0
#     answer = input('ERROR missing columns. Do you want to continue? (yes or no):')  #Creating answer menu
#     if answer.lower().startswith("yes"):    #Criteria if answer = yes
#         print("Proceeding without columns")
#         Columns = 0
#     elif answer.lower().startswith("no"):   #Criteria if answer = no
#         print("Executing process. Please update source model")


# if len(Beams) <= 0:
#     answer = input('ERROR missing beams. Do you want to continue? (yes or no):')
#     if answer.lower().startswith("yes"):
#         print("Proceeding without beams")
#         Beams = 0
#     elif answer.lower().startswith("no"):
#         print("Executing process. Please update source model")


#########################################################
##### Extracting properties for structural elements #####
#########################################################

beam_num = []
num = 0

name_beam = []
ref_level_beam = []
length_beam = []
material_beam = []
coord_beam = []


for beam in Beams:
#
    num = num+1
    b_num = "beam"+str(num)
    beam_num.append(b_num)

    #Extracting Reference level
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Constraints":
                for property in property_set.HasProperties:
                    if property.Name == "Reference Level":
                        ref_level = str(property.NominalValue.wrappedValue)
                        ref_level_beam.append(ref_level)

    #Extracting Length
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Dimensions":
                for property in property_set.HasProperties:
                    if property.Name == "Length":
                        length = str(round(property.NominalValue.wrappedValue, 2))
                        length_beam.append(length)

    #Extracting Material
    for definition in beam.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition
            if property_set.Name == "PSet_Revit_Materials and Finishes":
                for property in property_set.HasProperties:
                    if property.Name == "Beam Material":
                        material = str(property.NominalValue.wrappedValue)
                        material_beam.append(material)

    #Extracting Object coordinates
    xval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[0],3)
    yval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[1],3)
    zval = round(beam.ObjectPlacement.RelativePlacement.Location.Coordinates[2],3)
    coord = str([xval,yval,zval])
    coord_beam.append(coord)

print(beam_num)
print(ref_level_beam)
print(length_beam)
print(coord_beam)







