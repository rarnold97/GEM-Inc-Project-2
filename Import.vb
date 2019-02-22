'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Sub Import_Data()

Dim choice As String
Dim frm As New ReportChoiceFrm
Dim ButtonWS As Worksheet
'open a form too see what report type to use
frm.Show

If frm.Cancelled = True Then
    MsgBox "The UserForm was cancelled"
    Exit Sub
End If
On Error GoTo ErrorLine

choice = frm.Print_Choice

Set frm = Nothing

Set ButtonWS = Sheets("Report Macro Buttons")
'set the first cell to the report type so the user knows what they selected at all times
With ButtonWS
    .Cells(1, 1).Value = "Report Type:"
    .Cells(1, 2).Value = choice
End With

Dim Ws As Worksheet
Application.DisplayAlerts = False
'loop through each sheet and delete it if it is not sheet1.  If Sheet 1 is named "Sheet1" rename it
For Each xWs In Application.ActiveWorkbook.Worksheets
    If xWs.Name <> "Sheet1" Then
        If xWs.Name = "Information Imported from CGC" Then
            xWs.Delete
        End If
    Else
        xWs.Name = "Report Macro Buttons"
    End If
Next

Set Ws = ActiveWorkbook.Sheets.Add 'Add a new sheet to the workbook

With Ws
.Name = "Information Imported from CGC" 'Name the sheet

File = Select_File() 'prompt user ot select a csv data file

Open File For Input As #1 'open user selected file
Ctr = 0 ' Counter for rows of file
'Start a while loop
Do
    Line Input #1, Data 'read in the data line by line and add a new row for each line
    Ctr = Ctr + 1
    .Cells(Ctr, 1).Value = Data 'Populate the rows of the sheet with each line of the csv file
Loop While EOF(1) = False 'Loop until the end of the file is reached
Close #1 'Close the file that was read in

'ADD A SWITCH CASE HERE
'Modify Column Header Properties
Select Case choice
    'this formats the data differently depending on what report type the user chose
    Case "Budget Costs and Labor Hours"
        .Cells(1, 1).Resize(Ctr, 1).TextToColumns Destination:=Range("A1"), _
            DataType:=xlDelimited, Comma:=True, _
            FieldInfo:=Array(Array(1, xlGeneralFormat), Array(2, xlGeneralFormat), Array(3, xlGeneralFormat), Array(4, xlGeneralFormat), _
            Array(5, xlGeneralFormat), Array(6, xlGeneralFormat), Array(7, xlGeneralFormat), Array(8, xlGeneralFormat), _
            Array(9, xlGeneralFormat), Array(10, xlGeneralFormat), Array(11, xlGeneralFormat), Array(12, xlGeneralFormat), _
            Array(13, xlGeneralFormat), Array(14, xlGeneralFormat), Array(15, xlGeneralFormat), Array(16, xlGeneralFormat), _
            Array(17, xlGeneralFormat))
    Case "Crew Size and Labor Hours by Labor Type"
        .Cells(1, 1).Resize(Ctr, 1).TextToColumns Destination:=Range("A1"), _
            DataType:=xlDelimited, Comma:=True, _
            FieldInfo:=Array(Array(1, xlGeneralFormat), Array(2, xlGeneralFormat), Array(3, xlGeneralFormat), Array(4, xlGeneralFormat), _
            Array(5, xlGeneralFormat), Array(6, xlGeneralFormat), Array(7, xlGeneralFormat), Array(8, xlGeneralFormat), _
            Array(9, xlGeneralFormat), Array(10, xlMDYFormat), Array(11, xlGeneralFormat), Array(12, xlGeneralFormat), _
            Array(13, xlGeneralFormat), Array(14, xlGeneralFormat), Array(15, xlGeneralFormat), Array(16, xlGeneralFormat), _
            Array(17, xlGeneralFormat), Array(18, xlGeneralFormat), Array(19, xlGeneralFormat))
    Case "Crew Size and Labor Hours by Trade"
        .Cells(1, 1).Resize(Ctr, 1).TextToColumns Destination:=Range("A1"), _
            DataType:=xlDelimited, Comma:=True, _
            FieldInfo:=Array(Array(1, xlGeneralFormat), Array(2, xlGeneralFormat), Array(3, xlGeneralFormat), Array(4, xlGeneralFormat), _
            Array(5, xlGeneralFormat), Array(6, xlGeneralFormat), Array(7, xlGeneralFormat), Array(8, xlGeneralFormat), _
            Array(9, xlGeneralFormat), Array(10, xlMDYFormat), Array(11, xlGeneralFormat), Array(12, xlGeneralFormat), _
            Array(13, xlGeneralFormat), Array(14, xlGeneralFormat), Array(15, xlGeneralFormat), Array(16, xlGeneralFormat), _
            Array(17, xlGeneralFormat), Array(18, xlGeneralFormat), Array(19, xlGeneralFormat))
End Select

End With

Workbooks(Application.ActiveWorkbook.Name).Sheets("Report Macro Buttons").Activate  'Brings user to the report buttons sheet

Exit Sub

ErrorLine:
    MsgBox Err.Description

End Sub
