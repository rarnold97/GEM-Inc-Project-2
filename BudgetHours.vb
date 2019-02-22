'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Sub MainBudgetHours()

Dim Ws As Worksheet
Dim FinalRow As Integer
Dim AllCostCodes, AllTypes, AllDates As Variant
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
Dim CumHours, CumCost As Variant
Dim Codes As Collection
Dim SingleCode As String
Dim CostTypeDef As String
Dim GraphShtName As String
Dim SheetFrm As New AddGraphFrm
Dim GraphChoice As String
Dim ShtSelectFrm As New PlotSheetSelect
Dim Day As Boolean
Dim choice As Long

Day = False
'Turn off screen updating to make the code run faster
Application.ScreenUpdating = False


SheetFrm.Show 'asks if user wants to add a sheet or use their own for graphing purposes
'exit the code if the user cancells
If SheetFrm.Cancelled = True Then
    MsgBox "Userform was cancelled"
    Exit Sub
End If
'extract if the user choice
GraphChoice = SheetFrm.Return_Choice
If GraphChoice = "YES" Then 'user wants to add their own
    GraphShtName = AddSheet 'Prompt user to name a new sheet for graphs and then store the name for the plotting function to be used later
    If GraphShtName = "NONE" Then Exit Sub ' there are no appropriate places to graph, exit and try again
    
ElseIf GraphChoice = "NO" Then 'user opts to use their own sheet. must be blank
    If ShtSelectFrm.NoListProp = True Then Exit Sub ' exit because there are no available sheets
    ShtSelectFrm.Show ' show the form that picks a sheet that is not the command button sheet or the data sheet
    'exit if the user cancells
    If ShtSelectFrm.Cancelled = True Then
        MsgBox "The UserForm was cancelled"
        Exit Sub
    End If
    'extract the name of the sheet the user selects
    GraphShtName = ShtSelectFrm.Set_Sheet
    'kill variable for computational speed and to avoid crashing
    Unload ShtSelectFrm
    Set ShtSelectFrm = Nothing

End If
'kill form
Unload SheetFrm
Set SheetFrm = Nothing
'set active sheet
Set Ws = Sheets("Information Imported from CGC")

With Ws
'record final row of data
FinalRow = .Cells(Rows.Count, 1).End(xlUp).Row

choice = 1 'a choice of 1 means we are dealing with the budget type of report
frm.SetReportFlag = choice

'****************************************************************************************************
'Note, this can be changed to a function that seeks a header string rather than hard coding a column
'assuming that CGC excel exports stay the same
frm.Column1 = 7 'Cost Code Column header identifier
frm.Column2 = 8 'Cost Code Type Type Header identifier
'*****************************************************************************************************
frm.Show ' show the GUI to the user
'exit if user cancells
If frm.Cancelled = True Then
    MsgBox "The UserForm was cancelled"
    Exit Sub
End If
'set the selections based on the form stored in a collection variable
Set Codes = New Collection
Set Codes = frm.ReturnSelections
'kill the form
Unload frm
Set frm = Nothing

'Making sure the user actually selected something. If not , the code terminates to prevent excel from crashing
If Codes.Count = 0 Then
    Exit Sub
End If

' entry is the selection from the user. may contain numbers or strings, declared as a variant type
Dim entry As Variant
'loop throught the selections
For Each entry In Codes
    Set UniqueDates = CreateObject("Scripting.Dictionary") ' create a dictionary to take advantage of the fact that dictionaries dont duplicate keys
    'set the cost code and identifer to a var
    SingleCode = entry
    'create a temporary var
    Dim TempCode() As String
    'split the variable in to two parts the cost code and the identifier
    TempCode = Split(SingleCode, ",")  ' See Function SplitCode for clarification
    Cost_Code = TempCode(0) 'First part of the string is the cost code
    Cost_Type = TempCode(1) 'Second part of the string is the code identifier
    'CostTypeDef = CostTypeDict(Cost_Type)
    Dim TypeDict As Object 'create an identifier definition dictionary
    Set TypeDict = DefineTypes(choice)
        'loop through the data in the sheet
        For i = 2 To FinalRow
            'set check variables to compare to what is in the stored selections
            CstCodeChck = Trim(.Cells(i, 7).Value)
            TypeCheck = TypeDict(.Cells(i, 8).Value)
            DateString = .Cells(i, 9).Value
            DateCheck = ToDate(DateString)
            'enter only if the value was selected by the user
            If CstCodeChck = Cost_Code And TypeCheck = Cost_Type Then
                'add hours and costs to the right places withing the class instances that I created
                If UniqueDates.Exists(DateCheck) Then
                    Set SingleDate = UniqueDates.Item(DateCheck) ' call out an object instance within a dictionary collection
                    'add values from spreadsheet
                    SingleDate.Add_Hours = .Cells(i, 16).Value 'hard coded column headers for labor hours and cost. Subject to change
                    SingleDate.Add_Budget = .Cells(i, 17).Value
                    
                    Set UniqueDates(DateCheck) = SingleDate 'save the updated class instance
                Else 'if it is the first time encountering the selection, add it to the dictionary
                    Set SingleDate = New dates
                    
                    SingleDate.Set_Name = DateCheck
                    SingleDate.Add_Hours = .Cells(i, 16).Value
                    SingleDate.Add_Budget = .Cells(i, 17).Value
                    
                    UniqueDates.Add DateCheck, SingleDate
                End If
            End If
            
            
        Next i
    'sort the values
    Set UniqueDates = SortDict(UniqueDates)

    Dim dates As Collection
    Dim budgets As Collection
    Dim Hours As Collection
    Dim key As Variant
    ReDim CumHours(1 To UniqueDates.Count) As Variant
    ReDim CumCost(1 To UniqueDates.Count) As Variant
    'This takes advantage of the fact that collections are static
    Set dates = New Collection
    Set budgets = New Collection
    Set Hours = New Collection
    
    For Each key In UniqueDates.Keys
        Set SingleDate = UniqueDates.Item(key)
        
        dates.Add SingleDate.Print_Name
        budgets.Add SingleDate.Print_Budget
        Hours.Add SingleDate.Print_TotalHours
        
    Next key
    
    Dim j As Integer
    j = 2
    ' take all the values and make the cumulative discrete functions
    CumHours(1) = Hours(Hours.Count)
    
    For i = Hours.Count - 1 To 1 Step -1
        CumHours(j) = Hours(i) + CumHours(j - 1)
        j = j + 1
    Next i
    
    Set Hours = New Collection
    
    CumHours = ReverseArray(CumHours)
    
    For Each Item In CumHours
        Hours.Add Item
    Next Item
    
    j = 2
    
    CumCost(1) = budgets(budgets.Count)
    
    For i = budgets.Count - 1 To 1 Step -1
        CumCost(j) = CumCost(j - 1) + budgets(i)
        j = j + 1
    Next i
    
    CumCost = ReverseArray(CumCost) ' reverse the array because excel is werid and plots future dates before past dates.......I dunno man ??
    
    Set budgets = New Collection
    
    For Each Item In CumCost
        budgets.Add Item
    Next Item
    
    Set UniqueDates = Nothing
    
    'Plot stuff here
    PlotData dates, budgets, Hours, Cost_Code, Cost_Type, "Cost ($)", "Hrs", GraphShtName, Day

Next entry

'Vertically Arrange all the graphs exported
SortGraphs (GraphShtName)

Charts2PDF GraphShtName
'PlotData dates, budgets, Hours, Cost_Code, Cost_Type, "Cost ($)", "Hrs"
End With

End Sub
