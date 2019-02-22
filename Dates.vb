'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Private Name As Date
Private TotalHrs As Long
Private Budget As Currency
Private Hours As Long
Private CrewSize As Integer
Private Crew As Object


Private Sub Class_Initialize() 'Code to execute once form variable is dimentionalized in the main code
    'set initial values to zero
    Hours = 0
    Budget = 0
    CrewSize = 0
    'make a dictionary of the crew to tell how many unique workers there are.  Dictionary objects cant duplicate keys, which are the worker names from the data sheet
    Set Crew = CreateObject("Scripting.Dictionary")
    
End Sub

Property Get Print_Name() As Date ' print the name of the instance of the class, which is the date in this case
    Print_Name = Name
End Property

Property Let Set_Name(Input_Name As Date) ' let the main code set the date
    Name = Input_Name
End Property

Property Let Add_Hours(Value As Long) ' add hours from the data sheet to the running total
    Hours = Hours + Value
End Property

Property Let Add_Budget(Value As Currency) ' add cost to the running total according to the data sheet.
    Budget = Budget + Value
End Property

Property Get Print_TotalHours() As Long ' print the total hours to the main code
    Print_TotalHours = Hours
End Property

Property Get Print_Budget() As Currency '  print the total budget to the main code
    Print_Budget = Budget
End Property

Property Let Add_CrewSize(val As Integer) ' add to the crew size according to the data sheet
    CrewSize = CrewSize + val
End Property

Property Get Print_CrewSize() As Integer ' print the crew size to the main code
    Print_CrewSize = CrewSize
End Property

Public Function CheckName(key As Variant)
    If Not Crew.Exists(key) Then 'see if the scanned member exists
        Crew.Add key, 1 'add the member to the dictionary
        CrewSize = CrewSize + 1 ' add to the crew size, which will be initially 0 in this case
    End If
End Function

