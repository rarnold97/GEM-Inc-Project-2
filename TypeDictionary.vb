'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Public Function DefineTypes(choice As Long) As Object
'Defines the Cost Code Types.  If I did not account for all of these, add to the dictionary the appropriate entry
Set DefineTypes = CreateObject("Scripting.Dictionary") 'make a dictionary to define the cost codes and labor types
If choice = 1 Then ' choice is a flag that directs the definitions depending on the report type
    With DefineTypes
        .Add "Y", "Non-Craft Labor"
        .Add "L", "Craft Labor"
        .Add "F", "Fab Shop Craft Labor"
        .Add "O", "Other"
        .Add "M", "Taxable Material"
        .Add "N", "NonTaxable"
        .Add "S", "SubContracts"
        .Add "E", "Equipment Rental Outside"
        .Add "G", "Equipment Internal"
        .Add "I", "Incident Labor" 'Modify this not sure what it is
    
    End With

ElseIf choice = 2 Then

    'Enter definitions for class code here
    With DefineTypes
        .Add 4, "Alpha"
        .Add 16, "Beta"
        .Add 19, "Gamma"
        .Add 900, "Delta"
    End With

End If

End Function
