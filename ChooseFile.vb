'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Public Function Select_File() As String
'Ask which file to copy.  In this case, we are interested in the csv file that contains the welds,
'NDE, project, etc.
X = Application.GetOpenFilename(FileFilter:="CSV Files (*.csv),*.csv", _
Title:="Choose File to Copy", MultiSelect:=False) ' Makes sure that the file we are looking for is a .csv, which is the information the database provides

If X = "False" Then  'If the file cannot be opened then exit the function and display message
    MsgBox "Failed to open file, please try again"
    Exit Function
End If

MsgBox "You selected " & X  'displays to the user what file they selected

Select_File = X  'Returns the file address for other modules to manipulate

End Function



