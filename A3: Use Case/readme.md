3A: Analyse use case (JBM)  

Goal: Goal of the tool / workflow in one sentence. i.e. to support the user to calculate the total cost of the project. 

To calculate on carrying elements in the ‘Design review’ stage 

 

Model Use (Bim Uses): Please refer initially to the Mapping BIM uses, use cases and processes section in this document. 

Se PDF “BIM Use assignment” 

3B: Propose a (design for a) tool / workflow (OHH) 

Process: model the process diagram from your use case in BPMN.io please remember to save the .bpmn file and you can save a .svg file that you can insert into your report. 
<IMG src="IMG/A3 BPMN Group 14.svg">

description of the process of your tool / workflow. 

An IFC model gets loaded into the tool, it then searches after the load carrying elements in the model, and checks if there is any.  
If there isn't any the tool will stop. If there is some then it will extract the necessary information, like the geometry, span and material. Then it will calculate the rest of the necessary properties. A load case will be loaded into the tool and then a structural analysis will take place, to see if the model lives up to the standards. If not, it will tell where there is a problem. If everything is fine, it will print a structural documentation.  

 

3C: Information Exchange (VR) 

Information Exchange: Fill out the excel template with the information for your planned tool / workflow. For this you will need access to the excel, and the Dikon document to help you specify the LOD (LOR,LOG,LOI) for each element you need for you tool / workflow. This can get confusing - don’t worry we can help 😊 

Design review, angiv detalje grad af relevante elementer. 

IFC: Describe the IFC entities and properties for each of the elements you identified in your information exchange. Describe the data that you need to: 

Find in the IFC 

To extract the relevant data for calculating if every structural element in the building comply with the given loads, the following IFC entities and properties have to be found: 

Amount of structural elements 

Material properties: the kind of material (concrete, steel, wood etc.), Young’s modulus, compressive and tensile strength. 

Geometry: length, width and height, cross section dimensions. 

Location: x, y and z location in the building, storey. 

Find in an external sources i.e. BR18 

For the calculations of the structural elements to be useful, they have to comply with the norms in BR18 and Eurocode.  
For BR18 an element must live up to the requirements of e.g. fire safety. Structural elements must be placed in a way that there can be enough emergency exits and they must be constructed in way which does not catch fire. 
For Eurocode the right formulas must be used and the requirements for the relevant formulas must be fulfilled. E.g. the deformation of a steel beam must not exceed a specific value in order to allow people to be comfortable. 

 

Based on assumptions (useful when we don't have the 'real' data that we need for our tool) 

3D: Value what is the potential improvement offered by this tool? (JMH) 

This is the common question when developing tools and processes as an intrapreneur in a company. You should consider the business and societal value of this tool – does it save time to the company, does it make employees happier / more productive? Could it reduce material use in society? 

Describe the business value (How does it create value for your business / employer)  
 
For architect businesses it could end up saving a lot of time. If the tool were implemented in every IFC file or similar, it would point out immediately to the designer where there are flaws. This could therefore also save some person-hours from the engineer, which is expensive for the building project.  
It could be beneficial for businesses that are renovating a lot by using the tool on older buildings. (If the drawings are suited for it of course). The tool would in this scenario make it quicker to see if the relevant capacities is fulfilled when for instance removing a wall or similar.  

Describe the societal value (How does it make the world better) 

The Societal value could for instance be the fact that the tool would essentially point out unacceptable dimensions for the carrying elements. This would give another level of security for the construction of new buildings.   

N.B. If it doesn't do either of these things (ideally it should do both - don't do it!!) 

3E: Delivery (OHH) 

In this assignment we will focus on the input data that you need for your final tool / workflow.  

9. Your tool/workflow: Description of how your tool / workflow would solve the use case 

The tool loads in a IFC model and a load case put on the building. Then it will analyze the structural elements of the model and see if it lives up to the standards for a given load case. Then it should make a structural documentation for the building. But also make it easy for the engineers to extract important information about the structural elements, so the tools calculations easily could be checked.  

10. Delivery: Description of how you would make the tool / workflow - what steps would you go through?  

Our tool takes an IFC model, extracts the load carrying part of the model, it could be beams, columns, walls, etc. Then it collects the important information a structural engineer needs to know, like the material, cross section, length, etc. Then it will calculate the rest of the needed information that the model itself doesn’t contain.  Then a load case will be imported into the tool, so that it can calculate if the structure lives up to the standards. If something does not live up to the standards then the tool tells it, if everything is good, then it will print a structural documentation. 

 

 
