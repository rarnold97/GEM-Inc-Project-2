'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Sub Main_CrewHrsByType()

Dim Ws As Worksheet
Dim FinalRow As Integer
Dim AllCostTypes, AllTypes, AllDates As Variant
Dim i As Integer
Dim UniqueDates As Object
Dim SingleDate As dates
Dim frm As New FilterForm
Dim Cost_Code As String
Dim Cost_Type As String
Dim CstCodeChck As String
Dim TypeCheck As String
Dim DateString As String
Dim DateCheck As Date
Dim Size As Integer
Dim Person As String
Dim CrewCount As Integer
Dim Label1, Label2 As String
Dim CumHours As Variant
Dim LabelFrm As New LabelType
Dim Label As String
Dim Description As String
Dim TypeDefinitions As Object
Dim Codes As Collection
Dim SingleCode As String
Dim GraphShtName As String
'Dim CostDef As String
Dim SheetFrm As New AddGraphFrm
Dim GraphChoice As String
Dim ShtSelectFrm As New PlotSheetSelect
Dim Day As Boolean
Dim choice As Long

'Turn off screen updating to make the code run faster
Application.ScreenUpdating = False
'show user form to see if user wants to add a new sheet to add the report to
SheetFrm.Show
'exit the sub routine if the user cancelled
If SheetFrm.Cancelled = True Then
    MsgBox "Userform was cancelled"
    Exit Sub
End If
'extract the user choice
GraphChoice = SheetFrm.Return_Choice
If GraphChoice = "YES" Then 'user wants to add a new sheet for plots
    GraphShtName = AddSheet 'Prompt user to name a new sheet for graphs and then store the name for the plotting function to be used later
    If GraphShtName = "NONE" Then Exit Sub 'if there are no available sheets exit the sub routine
    
ElseIf GraphChoice = "NO" Then ' user wants to select an existing sheet
    If ShtSelectFrm.NoListProp = True Then Exit Sub 'no list is available, invalid sheets are shown
    ShtSelectFrm.Show 'bring up user form
    'note only blank sheets can be used and sheetes that are not the macro button sheet and the data sheet
    If ShtSelectFrm.Cancelled = True Then       'user cancelled the form exit the sub routine
        MsgBox "The UserForm was cancelled"
        Exit Sub
    End If
    'have user select the sheet to graph to after this
    GraphShtName = ShtSelectFrm.Set_Sheet
    'set to nothing to conserve memory
    Unload ShtSelectFrm
    Set ShtSelectFrm = Nothing

End If
'set form to nothing to conserve memory
Unload SheetFrm
Set SheetFrm = Nothing
'set worksheet to the data sheet to do some data minning
Set Ws = Sheets("Information Imported from CGC")
'used for extracting the codes the user selected
Set Codes = New Collection

'Set TypeDefinitions = DefineTypes

With Ws
'count the number of used rows
FinalRow = .Cells(Rows.Count, 1).End(xlUp).Row
'this is a flag for the dictionary that defines the cost code identifiers
choice = 1
'pass in the flag to the form before pulling it up
frm.SetReportFlag = choice
'can be changed depending on data sheet being used
'''''''''''''''''''''''''''''''''''''''''''''''''''
frm.Column1 = 8 'Cost Code Column header identifier
frm.Column2 = 9 'Cost Type Header identifier
'''''''''''''''''''''''''''''''''''''''''''''''''''

'pull up to user form
frm.Show

If frm.Cancelled = True Then 'exit the sub routine if the user cancells
    MsgBox "The UserForm was cancelled"
    Exit Sub
End If
'record the choices the user made
Set Codes = frm.ReturnSelections
'set the form to nothing to conserve memory
Unload frm
Set frm = Nothing
'no codes showed up exit the sub routine
If Codes.Count = 0 Then
    Exit Sub
End If
'pull up the user form
LabelFrm.Show
'user cancells exit the sub routine
If LabelFrm.Cancelled = True Then
    MsgBox "The UserForm was cancelled"
    Exit Sub
End If
'record the label type as either week ending date or day for the graph x-axis labels
Label = LabelFrm.Print_Choice
'this is a flag for the plotting function that does not apply to the cost/hour routine, but applies to this and CrewHrs subs
If Label = "Day" Or Label = "Week Ending Date" Then Day = True
'set the form to nothing to conserve memory
Unload LabelFrm
Set LabelFrm = Nothing
'if there is no label stored, exit the sub routine
If IsEmpty(Label) Then
    Exit Sub
End If
'variable used as a temporary value to examine the cost code and identifer the user selects
Dim entry As Variant
'scan through all the user selections
For Each entry In Codes
    Set UniqueDates = CreateObject("Scripting.Dictionary") 'make a dictionary to store unique values of dates only
    
    SingleCode = entry
    
    Dim TempCode() As String
    'parse the selection in to a cost code and an identifier .  the comma comes from one of the forms
    TempCode = Split(SingleCode, ",")
    'record cost code and cost code identifier
    Cost_Code = TempCode(0)
    Cost_Type = TempCode(1)
    'CostDef = "Test Placeholder"
    Dim TypeDict As Object
    Set TypeDict = DefineTypes(choice) 'instantiate a dictionary that defines the cost code identifier
        'scan the data sheet and mine for the info we are interested in
        For i = 2 To FinalRow
            CstCodeChck = Trim(.Cells(i, 8).Value) 'record the cost code, identifier, and then the person to keep track of the number of people on a job
            TypeCheck = TypeDict(.Cells(i, 9).Value)
            Person = .Cells(i, 13).Value
            
            Select Case Label 'add dates depending on whether or not the week ending date or specific day was chosen
                Case "Week Ending Date"
                    DateCheck = .Cells(i, 10).Value
                Case "Day"
                    'Adjusting the date based on the number of days Entry, subtract days because datasheet specifies weekending date.  Can be added insteaed if needed
                    DateCheck = .Cells(i, 10).Value - .Cells(i, 11).Value 'change to plus if necessary, wasnt sure if week ending date meant when the week ends or starts
            End Select
            
            If CstCodeChck = Cost_Code And TypeCheck = Cost_Type Then ' this if statement checks if the type selected matches the value scanned in the data sheet
                
                If UniqueDates.Exists(DateCheck) Then 'add to an existing value
                    'set values from data sheet/add them
                    Set SingleDate = UniqueDates.Item(DateCheck)
                    
                    SingleDate.Add_Hours = .Cells(i, 17).Value + .Cells(i, 18).Value + .Cells(i, 19).Value
                    SingleDate.CheckName Person
                    'add back to the dictionary
                    Set UniqueDates(DateCheck) = SingleDate
                Else 'add another class instance
                    Set SingleDate = New dates
                    'set the new values
                    SingleDate.Set_Name = DateCheck
                    SingleDate.Add_Hours = .Cells(i, 17).Value + .Cells(i, 18).Value + .Cells(i, 19).Value
                    SingleDate.CheckName Person
                    'add back to the dictionary
                    UniqueDates.Add DateCheck, SingleDate
                End If
            End If
            
            
        Next i
    'sort the dictionary values in order , see the function for this
    Set UniqueDates = SortDict(UniqueDates)
    
    Dim dates As Collection
    Dim CrewSizes As Collection
    Dim Hours As Collection
    Dim key As Variant
    
    Set dates = New Collection
    Set CrewSizes = New Collection
    Set Hours = New Collection
    ReDim CumHours(1 To UniqueDates.Count) As Variant
    'extract out values to find cumulative sums
    For Each key In UniqueDates.Keys
        Set SingleDate = UniqueDates.Item(key)
        
        dates.Add SingleDate.Print_Name
        CrewSizes.Add SingleDate.Print_CrewSize
        Hours.Add SingleDate.Print_TotalHours
        
    Next key
    
    'find the cumulitive sum and go backwards in the loop because excel orders dates in arrays weird.  Likely not the most efficient way, but could be modified
    Dim j As Integer
    j = 2
    
    CumHours(1) = Hours(Hours.Count)
    
    For i = Hours.Count - 1 To 1 Step -1
        CumHours(j) = Hours(i) + CumHours(j - 1)
        j = j + 1
    Next i
    
    Set Hours = New Collection
    
    CumHours = ReverseArray(CumHours) ' I incorporated this because for some reason the higher cumulative hours are recorded at earlier dates. Has to do with For Each Order or Date Type
    
    For Each Item In CumHours
        Hours.Add Item
    Next Item
    
    Set UniqueDates = Nothing
    
    PlotData dates, CrewSizes, Hours, Cost_Code, Cost_Type, "Crew Size", "Hrs", GraphShtName, Day
    'PlotData dates, CrewSizes, Hours, Cost_Code, Cost_Type, "Crew Size Count", "Hrs"
Next entry

'sort the graphs and make them neater in the plot report sheet
SortGraphs (GraphShtName)
'export all graphs to separate pages in a pdf
Charts2PDF GraphShtName

End With

End Sub
