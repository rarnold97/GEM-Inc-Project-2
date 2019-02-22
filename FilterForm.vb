'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019
Private Sub UserForm_Activate()
    FillListNoDuplicates 'wait until the user activates the form in order to populate the list with no duplicates
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
'the remaining properties return the specified values to the main sub routines
Public Property Let Column1(val As Integer)
    col1 = val
End Property

Public Property Let Column2(val As Integer)
    col2 = val
End Property

Public Property Let Add_Size(val As Integer)
    CrewSize = CrewSize + 1
End Property

Public Property Get Print_CrewSize() As Integer
    Print_CrewSize = CrewSize
End Property

Public Property Get ReturnSelections() As Collection
    Set ReturnSelections = CodeSelections
End Property

Public Property Let SetReportFlag(val As Long)
    choice = val
End Property
