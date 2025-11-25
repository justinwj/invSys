invSys - inventory management in Excel


1. What a tally system is in invSys:
	- a small collection of UI buttons, forms, user interface enhancements that allow the user to rapidly enter received items into inventory, or if in shipping, rapidly enter outgoing boxes.
		- in invSys every outgoing item is a stored collection of inventory items. Those items are tracked by invSys and deducted from inventory when outgoing boxes are tallied in the shipping sheet
		- received items that need to be added to inventory can be rapidly added by selecting the list object, which opens item search, showing a list of items that are managed in invSys
			- adding items that are not tracked to invSys is easy. Items are added by adding them to the list object, 'invSys', in the INVENTORY MANAGEMENT sheet. invSys list object manages with the help of a few buttons all tracked inventory that received, used, made and/or shipped.

2. All entries should be logged:
	- Received, Used, Made and Shipped items and the collections of items (boxes) are all logged

3. Production management (not implemented yet) logs Production batches and recall codes. 
	- recipe management is separate from the items used to make the recipes
	- recipes are stored in list objects, items that will actually be used are selected from a limited hand picked list of items deemed acceptable e.g. for pure cane sugar both Zucar Mex pure cane sugar and C&H pure cane sugar are acceptable choices. this 'substitutions' list is list objects created as needed.
	- combining acceptable substitutions is another feature invSys needs e.g. 200lb Zucar Mex pure cane sugar + 175lb C&H pure cane sugar.
			- the recipe lists pure cane sugar as the ingredient, the substitutions list is where what is considered pure cane sugar is defined.
	- invSys Production management needs to support adding variants of a recipe. Sometimes acceptable ingredients cannot be found, so i variant recipe needs to be developed e.g. we can't get the powdered lemon or the 32 brix apple syrup, there are no acceptable substitutes, so we have to create a variant recipe using lemon juice and apple sauce.

4. Printing:
	- recall codes need to be generated for items made. They will be sent to probably ZebraDesigner Pro 3.
	- production sheets need to be designed that are generated on click when a batch is prepared for making.
	- assembly labels may need to be printed to keep batches organized.

5. VBA UI
    - i have been using FRM files to create forms and form controls offers buttons, but using Ai to write forms and buttons makes it easier for the Ai to know what going on in the program. this process is faster for iteration with an Ai. Mermaid diagrams can be written by the Ai to review form and other controls design. good inline commenting documenting all the parts keeps the system organized. form and control design code should probably be kept in the same module as the programming logic they use.

6. Dashboards, Power Pivot, Slicers, Power Query, KPI's, measures, diagrams, chart etc...
	- after VBA is used to achieve needed functionality, Excel powerful tools should be used to extend insight gathering in invSys
	- data connections between workbooks might be used.

7. Future releases Streamlit, .NET
	- Streamlit will likely be used to develop streamer front ends to send data to invSys, hardening the data ingestion process, but this would not be done in the first release.
	- .NET might be used in future releases