
# 3A: Analyse use case


### 1) *Goal of the tool / workflow*
To single out critical bearing elements for further structural analysis in the design face. 

### 2) *Model Use (Bim Uses)*

BIM Use: 
- Fasility element: 
    - Substructure


# 3B: Propose a (design for a) tool / workflow


### 3) *Process*
<IMG src="IMG_folder/A3 BPMN Group 14.svg">

### 4) *description of the process*
An IFC model gets loaded into the tool, it then searches after the load bearing elements in the model, and checks if there is any.  
If there isn't any the tool will stop. If there is some then it will extract the necessary information, like the geometry, placement, material etc. Then it will calculate the rest of the necessary properties. A load case will be loaded into the tool and then a structural analysis will take place, to see if the model lives up to the standards. If not, it will tell where there is a problem. If everything is fine, it will print a structural documentation.  


# 3C: Information Exchange 3


### 5) *Information Exchange*
See attached excel file.
 
### 6) *IFC*
Find in the IFC:
To extract the relevant data for calculating if every structural element in the building comply with the given loads, the following IFC entities and properties have to be found: 

Amount of structural elements 

Material properties: the kind of material (concrete, steel, wood etc.), Young’s modulus, compressive and tensile strength. 

Geometry: length, width and height, cross section dimensions. 

Location: x, y and z location in the building, storey. 

Find in an external sources i.e. BR18: 
For the calculations of the structural elements to be useful, they have to comply with the norms in BR18 and Eurocode.  
For BR18 an element must live up to the requirements of e.g. fire safety. Structural elements must be placed in a way that there can be enough emergency exits and they must be constructed in way which does not catch fire. 
For Eurocode the right formulas must be used and the requirements for the relevant formulas must be fulfilled. E.g. the deformation of a steel beam must not exceed a specific value in order to allow people to be comfortable. 

 

# 3D: Value what is the potential improvement offered by this tool?

 
### 7) *Business value*
For architect businesses it could end up saving a lot of time. If the tool were implemented in every IFC file or similar, it would point out immediately to the designer where there are flaws. This could therefore also save some person-hours from the engineer, which is expensive for the building project.  
It could be beneficial for businesses that are renovating a lot by using the tool on older buildings. (If the drawings are suited for it of course). The tool would in this scenario make it quicker to see if the relevant capacities is fulfilled when for instance removing a wall or similar.  

### 8) *Societal value*
The Societal value could for instance be the fact that the tool would essentially point out unacceptable dimensions for the carrying elements. This would give another level of security for the construction of new buildings.   


# 3E: Delivery


### 9) *Our tool/workflow*
The tool loads in a IFC model and a load case put on the building. Then it will analyze the structural elements of the model and see if it lives up to the standards for a given load case. Then it should make a structural documentation for the building. But also make it easy for the engineers to extract important information about the structural elements, so the tools calculations easily could be checked.  

### 10) *Delivery*
Our tool takes an IFC model, extracts the load carrying part of the model, it could be beams, columns, walls, etc. Then it collects the important information a structural engineer needs to know, like the material, cross section, length, etc. Then it will calculate the rest of the needed information that the model itself doesn’t contain.  Then a load case will be imported into the tool, so that it can calculate if the structure lives up to the standards. If something does not live up to the standards then the tool tells it, if everything is good, then it will print a structural documentation. 

 

 
