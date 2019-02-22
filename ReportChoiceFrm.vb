'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Private m_Cancelled As Boolean
Private ReportChoice As String

Private Sub CommandButton1_Click()
    ReportChoice = ComboBox1.Value 'record the users choice
    Hide
End Sub

Private Sub UserForm_Initialize()
    FillComboList 'populate the form options to select from
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
    'output the form choice
    Print_Choice = ReportChoice
End Property

Private Function FillComboList()
    'add the possible options to the report type to guide the main subroutines to the right calculations
    ComboBox1.AddItem "Budget Costs and Labor Hours"
    ComboBox1.AddItem "Crew Size and Labor Hours by Labor Type"
    ComboBox1.AddItem "Crew Size and Labor Hours by Trade"

End Function

