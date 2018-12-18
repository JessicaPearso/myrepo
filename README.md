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

Option Explicit
Private m_empList As String

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
 If ActiveCell.Row <= 1 Then
        Exit Sub
      End If
        If ActiveCell.Column = 2 And ActiveCell.Row > 1 Then
            If m_empList = "" Then
                m_empList = GetEmployeeList(ActiveCell, Sh.Name)
            End If
           MakeCombo Target, m_empList
        End If   
End Function

4.'Rest API call to get all the Employees of a Dept. Calling the GetEmployeeList function inside Workbook_SheetSelectionChange 

Private Function GetEmployeeList(Target As Range, Name As String) As String
Dim objRequest As Object
Dim jsonDictionary As New Dictionary
Dim TempJsonString As String, JsonString As String
Dim strResponse As String
Dim jsonItems As New Collection
Dim jsonObject As Object, item As Object
Dim myValidationStr As String
'get the all the employees of a dept
Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
objRequest.Open "POST", "http://localhost:8080/getAllEmployees"

    jsonDictionary("dept") = "IT"
    jsonItems.Add jsonDictionary
    JsonString = JsonConverter.ConvertToJson(ByVal jsonItems)
    'Send Request.
    objRequest.send JsonString
    'And we get this response
    strResponse = objRequest.responseText
   
    If Trim(strResponse & vbNullString) <> vbNullString Then
     Set jsonObject = JsonConverter.ParseJson(strResponse)
     Dim i As Long
     With Sheets("employeesheet")  'copy all the employees in employeesheet
        .Range("A1:A1000").ClearContents
        i = 1
        For Each item In jsonObject("EmployeeList")
            i = i + 1
            .Range("A" & i).Value = item("empName")
        Next item
     End With
     
    End If
    
    GetEmployeeList = "=employeesheet!A1:A" & i 'copy all the employees in employeesheet
End Function
5.Use the below link to create the - Data Validation Combo box Click
https://www.contextures.com/xlDataVal14.html#works
'Setting the Employee list in the dropdown
Sub MakeCombo(ByRef Target As Range, ByRef comboList As String)
Dim cboTemp As OLEObject
Dim ws As Worksheet
Dim Tgt As Range
Dim TgtMrg As Range
Dim c As Range
Dim TgtW As Double
Dim AddW As Long
Dim AddH As Long

Set ws = ActiveSheet
On Error Resume Next
'extra width to cover drop down arrow
AddW = 15
'extra height to cover cell
AddH = 5

If Target.Rows.Count > 1 Then GoTo exitHandler

Set Tgt = Target.Cells(1, 1)
Set TgtMrg = Tgt.MergeArea
On Error GoTo errHandler

  Set cboTemp = ws.OLEObjects("TempCombo")
    On Error Resume Next
  If cboTemp.Visible = True Then
    With cboTemp
      .Top = 10
      .Left = 10
      .ListFillRange = ""
      .LinkedCell = ""
      .Visible = False
      .Value = ""
    End With
  End If

  On Error GoTo errHandler

    With cboTemp
      .Visible = True
      .Left = Tgt.Left
      .Top = Tgt.Top
      .Width = Tgt.Width 
      .Height = Tgt.Height + AddH
      .ListFillRange = comboList
      .LinkedCell = Tgt.Address
    End With
    cboTemp.Activate
    Me.ActiveSheet.TempCombo.DropDown


exitHandler:
  Application.EnableEvents = True
  Application.ScreenUpdating = True
  Exit Sub
errHandler:
  Resume exitHandler

End Sub
