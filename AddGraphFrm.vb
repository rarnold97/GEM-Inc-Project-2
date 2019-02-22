'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019
Private m_Cancelled As Boolean
Private choice As String

Private Sub Cancel_button_Click()
    'Hide the Userform and set cancelled to true
    Hide
    m_Cancelled = True
End Sub

Private Sub CommandButton1_Click()
    choice = "YES"
    Hide
End Sub

Private Sub CommandButton2_Click()
    choice = "NO"
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

Public Property Get Return_Choice() As String
    Return_Choice = choice
End Property
