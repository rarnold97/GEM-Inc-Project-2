'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Private Sub clear_Click()
    Application.DisplayAlerts = False ' turn of the box that asks user if they are sure they want to delete the sheet. I am overiding this
    
    Dim xWs As Worksheet
    ' delete all worksheets except for the macro button sheet and the data sheet
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name <> "Information Imported from CGC" And xWs.Name <> "Report Macro Buttons" Then
            xWs.Delete
        End If
    Next
    MsgBox "Cleared Report Worksheets"
End Sub

Private Sub CommandButton2_Click()
    Import_Data 'calls the import function
End Sub


Private Sub CommandButton4_Click()
    'Documentation link: https://www.extendoffice.com/documents/excel/785-excel-save-export-sheet-as-new-workbook.html
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim xWs As Worksheet
    Dim xWb As Workbook
    Dim FolderName As String
    Application.ScreenUpdating = False
    Set xWb = Application.ThisWorkbook
    DateString = Format(Now, "yyyy-mm-dd")
    Dim frm As New SelectSheetFrm
    Dim SheetChoice As String
    'FolderName = xWb.Path '& "\" & xWb.Name & " " & DateString
    'MkDir FolderName
    
    On Error GoTo ErrHndlr
    'Set a variable to correspond to the report sheet, which is always the third sheet
    If frm.IsNoValues = True Then Exit Sub
    
    frm.Show

    If frm.Cancelled = True Then
        MsgBox "The UserForm was cancelled"
        Exit Sub
    End If
    
    SheetChoice = frm.Print_Choice ' extract the sheet the user selected
    
    Set xWs = Sheets(SheetChoice) 'set worksheet to selection
    
    xWs.Copy 'Copy contents of the sheet
    If val(Application.Version) < 12 Then 'Determine which extension to use
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        Select Case xWb.FileFormat 'see what excel format the active workbook is
            Case 51:
                FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If Application.ActiveWorkbook.HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56:
                FileExtStr = ".xls": FileFormatNum = 56
            Case Else:
                FileExtStr = ".xlsb": FileFormatNum = 50
        End Select
    End If
    xFile = xWb.Path & "\" & DateString & " " & xWs.Name & FileExtStr 'name the file
    Application.ActiveWorkbook.SaveAs xFile, FileFormat:=FileFormatNum 'save the file
    Application.ActiveWorkbook.Close False ' Make sure the program doesnt close itself
    
    MsgBox "You can find the Report file in: " & xWb.Path 'Display to the user where to find their report
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHndlr:
    MsgBox Err.Description 'display error

End Sub


Private Sub CommandButton6_Click() ' directs code to right report sub
    Dim ReportChoice As String
    Dim Ws As Worksheet
    Dim xWs As Worksheet
    Dim counter As Long
    
    counter = 0 ' control variable
    
    On Error GoTo ErrHndlr
    ' scan through worksheets, and if there is no data sheet, exit the code
    For Each xWs In Application.Worksheets
        If xWs.Name = "Information Imported from CGC" Then counter = counter + 1 'break the code
    Next
    
    If counter = 0 Then
        MsgBox "There is no data... Please import data by selecting 'Import New Data' button."
        Exit Sub
    End If
    
    Set Ws = Sheets("Report Macro Buttons") ' set sheet to buttons
    
    With Ws
        ReportChoice = .Cells(1, 2).Value 'retrieve what the report type is
    End With
    
    Select Case ReportChoice ' direct code to call sub based on report type.  There are three main subs
        Case "Budget Costs and Labor Hours"
            MainBudgetHours
        Case "Crew Size and Labor Hours by Labor Type"
            Main_CrewHrsByType
        Case "Crew Size and Labor Hours by Trade"
            Main_CrewHrs
        Case Else
            MsgBox "No Data To Analyze. Make sure to Import Data."
    End Select
    
    Exit Sub
    
ErrHndlr:
    MsgBox Err.Description
    
End Sub

