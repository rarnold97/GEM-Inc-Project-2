'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Public Function AddSheet() As String

Dim Control As Boolean
Dim Ws As Worksheet
Dim SheetName As String

SheetName = InputBox("Enter Sheet Name for Graphs") 'let user enter a graph name

If IsEmpty(SheetName) = True Then 'check if the sheet name is empty
    AddSheet = "NONE" 'return a flag to say there is no sheet and exit the function
    Exit Function
    
ElseIf SheetName = vbNullString Then 'check to see if it is a null string
    MsgBox "Entry Cancelled, Please Select Button or run macro again to restart" ' there is no string entered, exits code
    AddSheet = "NONE" 'flag for main code
    Exit Function
End If
Control = True 'set a loop control variable

Do While Control = True
    For Each xWs In Application.ActiveWorkbook.Worksheets 'scan through all worksheets
        If xWs.Name = SheetName Then
            MsgBox "Please Enter a different name, this one already exists." 'name already exists
            SheetName = InputBox("Enter Sheet Name for Graphs")
        Else
            Control = False 'break the loop
        End If
    Next
Loop

Set Ws = ActiveWorkbook.Sheets.Add ' add a new sheet
Ws.Name = SheetName 'name it to what user selects

AddSheet = SheetName 'return the name to the main subs

End Function
