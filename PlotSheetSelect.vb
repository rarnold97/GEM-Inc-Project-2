'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

'The purpose of this form is so that the user can select what worksheet to plot their graphs to. Gives them the flexibility to plot multiple times in one run of the program
Private m_Cancelled As Boolean
Private SheetNm As String
Private NoList As Boolean

Private Sub UserForm_Initialize()
    Dim Sh As Worksheet 'declare a sheet variable to scan through the worksheets in the workbook
    For Each Sh In ActiveWorkbook.Sheets 'check all the availble sheets
        If Sh.Name <> "Information Imported from CGC" And Sh.Name <> "Report Macro Buttons" Then
            If WorksheetFunction.CountA(Sh.UsedRange) = 0 And Sh.Shapes.Count = 0 Then 'check for blank worksheets
                'here, we dont want the data sheet and the macro button sheet to be plotted to, exclude these
                'Also exclude blank worksheets
                ListBox1.AddItem Sh.Name 'add the rest of the sheets to the forum.
            End If
        End If
    Next Sh
    
    If ListBox1.ListCount = 0 Then
        MsgBox "There are no other worksheets or there are no blank worksheets, please add a blank sheet to plot report data."
        NoList = True
    End If
End Sub

Private Sub CommandButton1_Click()
    'record the value that the user selects
    SheetNm = Sheets(ListBox1.Value).Name
    Hide

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) 'handles when the user clicks the "x" button of the form
    If CloseMode = vbFormControlMenu Then Cancel = True
    Hide
    m_Cancelled = True
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property

Public Property Get Set_Sheet() As String
    'return the name of the chosen sheet
    Set_Sheet = SheetNm
End Property

Public Property Get NoListProp() As Boolean
    NoListProp = NoList ' flag for the main function if there is no contents to the list box
End Property
