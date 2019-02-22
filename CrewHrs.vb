'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Sub Main_CrewHrs()

Dim Ws As Worksheet
Dim FinalRow As Integer
Dim AllCostCodes, AllClasses, AllDates As Variant
Dim i As Integer
Dim UniqueDates As Object
Dim SingleDate As dates
Dim frm As New FilterForm
Dim Cost_Code As String
Dim Class_Type As Variant
Dim Codes As Collection
Dim SingleCode As String
Dim CstCodeChck As String
Dim ClassCheck As String
Dim DateString As String
Dim DateCheck As Date
Dim Size As Integer
Dim Person As String
Dim CrewCount As Integer
Dim Label1, Label2 As String
Dim CumHours As Variant
Dim LabelFrm As New LabelType
Dim Label As String
Dim DayToAdd As Variant
Dim GraphShtName As String
'Dim ClassDef As String
Dim ShtSelectFrm As New PlotSheetSelect
Dim SheetFrm As New AddGraphFrm
Dim GraphChoice As String
Dim Day As Boolean
Dim choice As Long

'Turn off screen updating to make the code run faster
Application.ScreenUpdating = False

SheetFrm.Show

If SheetFrm.Cancelled = True Then
    MsgBox "Userform was cancelled"
    Exit Sub
End If

GraphChoice = SheetFrm.Return_Choice
If GraphChoice = "YES" Then
    GraphShtName = AddSheet 'Prompt user to name a new sheet for graphs and then store the name for the plotting function to be used later
    If GraphShtName = "NONE" Then Exit Sub
    
ElseIf GraphChoice = "NO" Then
    If ShtSelectFrm.NoListProp = True Then Exit Sub ' if there arent acceptable sheets exit the sub routine
    ShtSelectFrm.Show 'show user form
    
    If ShtSelectFrm.Cancelled = True Then 'exit if the user cancells the form
        MsgBox "The UserForm was cancelled"
        Exit Sub
    End If
    
    GraphShtName = ShtSelectFrm.Set_Sheet 'show the form that allowed the user to select a sheet to plot to
    
    Unload ShtSelectFrm 'clear the form object
    Set ShtSelectFrm = Nothing

End If

Unload SheetFrm 'clear the form object
Set SheetFrm = Nothing

Set Ws = Sheets("Information Imported from CGC") 'set worksheet equal to data sheet to extract and manipulate

With Ws

FinalRow = .Cells(Rows.Count, 1).End(xlUp).Row 'count number of used rows

choice = 2 'used for setting up list options in user forum for selecting cost codes and identifiers.
frm.SetReportFlag = choice
'change this if new .csv is used
''''''''''''''''''''''''''''''''''''''''''''''''''''
frm.Column1 = 8 'Cost Code Column header identifier
frm.Column2 = 14 'Class Type Header identifier
''''''''''''''''''''''''''''''''''''''''''''''''''''

frm.Show

If frm.Cancelled = True Then 'if the user cancells the form exit the sub routine
    MsgBox "The UserForm was cancelled"
    Exit Sub
End If

Set Codes = New Collection ' set a collection that will have the selected cost codes from the form

Set Codes = frm.ReturnSelections

Unload frm 'clear object variable
Set frm = Nothing

If Codes.Count = 0 Then 'exit the sub routine if there are no codes displayed
    Exit Sub
End If
'Cost_Code = Trim(frm.CstCode)
'Class_Type = frm.CstCodeTyp

LabelFrm.Show 'show a form that labels plots by day or by week ending date

If LabelFrm.Cancelled = True Then 'ext the sub routine if the user cancells the form
    MsgBox "The UserForm was cancelled"
    Exit Sub
End If

Label = LabelFrm.Print_Choice 'extract user choice
If Label = "Day" Or Label = "Week Ending Date" Then Day = True 'used in a function later as a flag
Unload LabelFrm
Set LabelFrm = Nothing
'If block will exit the code if the user does not select something.  this prevents unwanted crashes .  User can just rerun macro
If IsEmpty(Label) Then 'if there is no label exit the sub routine
    Exit Sub
End If

Dim entry As Variant 'variable that temporarily extracts row element from data sheet

For Each entry In Codes
    Set UniqueDates = CreateObject("Scripting.Dictionary") 'create a dictionary for uniqueness properties
    
    SingleCode = entry
    
    Dim TempCode() As String
    
    TempCode = Split(SingleCode, ",") 'See Function SplitCode for clarification
    Cost_Code = TempCode(0) ' first part of string is cost code
    Class_Type = TempCode(1) ' Second part of string is the class type assuming this is the correctly selected report type by the user
    'ClassDef = "Test Placeholder"
    Dim TypeDict As Object
    Set TypeDict = DefineTypes(choice)
    
        For i = 2 To FinalRow 'Loop through all the data provided
            CstCodeChck = Trim(.Cells(i, 8).Value) ' Record the cost code of indicated row
            ClassCheck = TypeDict(.Cells(i, 14).Value)  ' Convert To a String So that it matches output from Splitting function
            Person = .Cells(i, 13).Value ' Record the name of the person that worked on the job in the database
            
            Select Case Label ' Direct code to the graph label selected by the user
                Case "Week Ending Date" ' present by week
                    DateCheck = .Cells(i, 10).Value
                Case "Day" 'present by each day
                    'Adjusting the date based on the number of days Entry, subtract days because datasheet specifies weekending date.  Can be added insteaed if needed
                    DateCheck = .Cells(i, 10).Value - .Cells(i, 11).Value
            End Select
            
            If CstCodeChck = Cost_Code And ClassCheck = Class_Type Then
                
                If UniqueDates.Exists(DateCheck) Then ' if the date already exists
                    Set SingleDate = UniqueDates.Item(DateCheck)
                    'add the data to the class instance
                    SingleDate.Add_Hours = .Cells(i, 17).Value + .Cells(i, 18).Value + .Cells(i, 19).Value
                    SingleDate.CheckName Person
                    
                    Set UniqueDates(DateCheck) = SingleDate
                Else 'add a new data
                    Set SingleDate = New dates
                    'add the data to the class instance
                    SingleDate.Set_Name = DateCheck
                    SingleDate.Add_Hours = .Cells(i, 17).Value + .Cells(i, 18).Value + .Cells(i, 19).Value
                    SingleDate.CheckName Person
                    
                    UniqueDates.Add DateCheck, SingleDate
                End If
            End If
        Next i
        
    'Sort the Dictionary by date
    Set UniqueDates = SortDict(UniqueDates)
    
    Dim dates As Collection
    Dim CrewSizes As Collection
    Dim Hours As Collection
    Dim key As Variant
    
    Set dates = New Collection
    Set CrewSizes = New Collection
    Set Hours = New Collection
    ReDim CumHours(1 To UniqueDates.Count) As Variant
    
    For Each key In UniqueDates.Keys 'now adding all the data sets to one colletion used for plotting
        Set SingleDate = UniqueDates.Item(key)
        
        dates.Add SingleDate.Print_Name
        CrewSizes.Add SingleDate.Print_CrewSize
        Hours.Add SingleDate.Print_TotalHours
        
    Next key
    
    
    Dim j As Integer
    j = 2
    
    CumHours(1) = Hours(Hours.Count) ' define the first point sarting backwards to sort correctly
    'excel keeps plotting the dates backwards so the corresponding y-values need to be reversed in the cumulative sum
    'make a cumulative sum of the costs
    For i = Hours.Count - 1 To 1 Step -1
        CumHours(j) = Hours(i) + CumHours(j - 1)
        j = j + 1
    Next i
    
    Set Hours = New Collection
    
    CumHours = ReverseArray(CumHours) ' I incorporated this because for some reason the higher cumulative hours are recorded at earlier dates. Has to do with For Each Order or Date Type
    
    For Each Item In CumHours 'take a cumulative sum for the hours
        Hours.Add Item
    Next Item
    
    Set UniqueDates = Nothing
    
    PlotData dates, CrewSizes, Hours, Cost_Code, Class_Type, "Crew Size", "Hrs", GraphShtName, Day
    'PlotData dates, CrewSizes, Hours, Cost_Code, Class_Type, "Crew Size Count", "Hrs"
    
Next entry
'sort the graphs in the plot report sheet
SortGraphs (GraphShtName)
'export each graph as a unique page in a pdf
Charts2PDF GraphShtName

End With

End Sub

