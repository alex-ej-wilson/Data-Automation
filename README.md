==========================================================================================================
For all programs:
	
	Python (3.11) will need to be installed - available on the windows store.

	The xlwings Python package will need to be added, this is done using the Windows PowerShell
	and using command "pip install xlwings"
	
	Corresponding Excel Workbooks will macros adding, instructions and code is contained within
	"VBA Excel Macros Instructions.txt". Using your selected macro will trigger the Python Script
	corresponding to the Workbook.
	
	The 'config.json' file contains editable data that is used by the programs. To edit the file
	open with Notepad or any other program that lets you edit plain text (no formatting)

==========================================================================================================
ItemDB_adder.py for the Item Database:

	Saving the file will trigger the macro, which will copy the item database sheet into the 
	example control document, as a hidden sheet. Drop down menus will also be created containing 
	the data as suggestions, so that unique values can also be entered. Keep in mind the way this
	program works means that only new control documents created will contain the updated item
	database.

==========================================================================================================
PLEASE NOTE THE FOLLOWING PROGRAMS WORK BASED ON THE MOST RECENTLY SAVED DATA, TO ENSURE THEY WORK AS
INTENDED, PLEASE SAVE THE WORKBOOK BEFORE USING THE MACRO.
----------------------------------------------------------------------------------------------------------
Schedule_to_folder.py for the Schedule:

	Triggering macro will generate a new project folder based for the last row in the Excel book on 
	the template, the folder will be named based on the Job Number, Customer, and Project values in 
	'Schedule.xlsm'. If the customer is in the "main_clients" list in 'config.json', the job folder 
	will be stored in their specific folder and only be named using the Job Number and Project 
	Reference, all others will be stored in the 'MISCELLANEOUS' folder and will be named using all
	three collected values. The file 'A-Control-Doc.xlsm' will have the Job Number, Customer Name,
	Project Reference and Delivery Date filled in.

----------------------------------------------------------------------------------------------------------
ControlDoc_to_Invoice.py for Control Documents:

	Triggering the macro will generate a pdf version of the Sales Invoice and Delivery Note sheet,
	based on the currently saved information. These files will be stored in a specified location,
	given in 'config.json'. Within the overall folder each note will be stored in a folder based 
	on its number in a range of 0 - 99, i.e. Job Number: 12356 would be stored in a File "12300 -
	12399".

==========================================================================================================
