'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Private m_Cancelled As Boolean
Private WSChoice As String
Private novalues As Boolean

Private Sub CommandButton1_Click()
    WSChoice = ComboBox1.Value 'record the users choice as a private variable
    Hide
End Sub

Private Sub UserForm_Initialize()
    FillComboList 'populate the list of options
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) 'handles when the user clicks the "x" button of the form
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    Hide
    
    m_Cancelled = True
    
End Sub


'Returns the cancelled value to the calling procedure
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property


Property Get Print_Choice() As String
    Print_Choice = WSChoice ' return the choice that the user selected
End Property

Private Function FillComboList()
    'loop through each sheet
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name <> "Report Macro Buttons" And xWs.Name <> "Information Imported from CGC" Then
            ComboBox1.AddItem xWs.Name ' we are not interested in the macro button and datasheet
        End If
    Next
        
    If ComboBox1.ListCount = 0 Then ' prompt user that they need a sheet that is not the data sheet or the macro button sheet
            MsgBox "No Valid Sheets, please add one and try again."
            novalues = True
    End If
End Function

Public Property Get IsNoValues() As Boolean
    IsNoValues = novalues 'flag that indicates to the main code that there are no values and to exit the sub
End Property











