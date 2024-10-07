# A2 - BIManalyst group 26

## A2a -About our group:
Our groups Pyhton level is about 2-3 (Neutral/Agree)

Our groups focus area is the sturctural (STR) aspect of the building, and we are analysts.

## A2b - Identify claim:
We have chosen building #XXXX

The group has decided to inspect the structural elements of the building. The first step is to gather data from the IFC regarding the number of beams, columns, walls and slabs throughout the structure. After this, we'll verify that the dimensions of these elements match those specified in the STR report.
Finally, we plan to make a cost estimate for the building's structural components.

## A2c - Use Case:
As explained in section A2b, our goal is to create a script that can automatically collect this data for us. This information must be verified during the design phase to ensure the structural elements are agreed upon. Additionally, it will allow contractors and other stakeholders to provide an estimated project cost.
The data required includes details about the dimensions and quantity of the structural elements, which are essential for both checking compliance and estimating costs during the design phase.

## A2d - Scope the use case:
BPMN and SVG files

## A2e - Tool idea:
In summary, the primary goal of our scripts is to derive an estimated cost for the structural elements of the building. By automating this process, we aim to provide crucial information that will be highly beneficial during the design phase.
For stakeholders, having early access to these cost estimates allows for more informed decision-making regarding budget allocation and potential adjustments to the design. Similarly, contractors can use this data to plan and prepare accurate bids and financial projections for the project. 
Overall, our scripts will streamline the cost estimation process, ensuring that both stakeholders and contractors have the information they need early on, which can help avoid costly revisions later in the project.


## A2f - Information Requirements
Each structural element is labeled in the IFC file, but we anticipate encountering some issues, as certain slabs in the file are incorrectly labeled as beams. In the previous A1 assignment, we successfully extracted information about the length of beams, so it should be feasible to retrieve similar data for this task by applying the same methods.

Regarding the materials used, we plan to use the report as a reference, since the model may lack some material information. However, the report provides clear details about the materials. This will ensure accuracy when calculating the cost estimate for the structural elements.

## A2g - Identify appropriate software license
- Blender
- Pyhton
- Bonsai
