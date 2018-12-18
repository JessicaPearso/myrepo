Populate dynamically dropdown in Excel with VBA using Rest API call

1. Add a reference to the “Microsoft Scripting Runtime”
  To reference this file, load the Visual Basic Editor (ALT+F11)
  Select Tools > References from the drop-down menu
  A listbox of available references will be displayed
  Tick the check-box next to 'Microsoft Scripting Runtime'
  
2. import the “JsonConverter.bas” file to your VBA project.
    - https://github.com/VBA-tools/VBA-JSON/issues/17
    
 Below code has been written in Workbook

3. Workbook_SheetSelectionChange event macro is the cell or range of cells that have just been selected
For second column in excel populating the drop down

4.'Rest API call to get all the Employees of a Dept. Calling the GetEmployeeList function inside Workbook_SheetSelectionChanged JsonString
   
5.Use the below link to create the - Data Validation Combo box Click
https://www.contextures.com/xlDataVal14.html#works
'Setting the Employee list in the dropdown
