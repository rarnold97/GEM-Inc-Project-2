'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019
Private LabelChoice As String

Private Sub CommandButton1_Click()
'Record what the user selects
    LabelChoice = ComboBox1.Value
    Hide
End Sub

Private Sub UserForm_Initialize()
    'populate the form with available options
    FillComboList
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
    'return what the user selected in the GUI
    Print_Choice = LabelChoice
End Property

Private Function FillComboList()
    'fill the combolist with the two available label options
    ComboBox1.AddItem "Week Ending Date"
    ComboBox1.AddItem "Day"
    
End Function


