'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Public Function PlotData(X As Collection, y1 As Collection, y2 As Collection, code As String, codedesc As Variant, y1Label As String, y2Label As String, _
WsName As String, Day As Boolean)

Dim Ws As Worksheet
Dim CH As Chart
Dim Ser1 As Series
Dim Ser2 As Series
Dim FinalRow As Integer
Dim rng1 As Range
Dim rng2 As Range
Dim i, j, k As Integer
Dim io, jo, ko As Integer
Dim val As Variant
Dim SheetName As String
'Dim Control As Boolean
Dim ChartTitle As String

ChartTitle = "Cost Code: " & code & "  " & "Labor Type/Trade: " & CStr(codedesc)

Set Ws = Sheets(WsName)
Ws.Select
With Ws


'Goes through the values on the x-axis, which in most cases is some sort of date
'Each item is entered in a spreadsheet cell
If IsEmpty(.Cells(1, 1).Value) = True Then
    io = 1 ' this if block is if we are on the first data set, so start in cell 1
    jo = 1
    ko = 1
Else 'this if block is if a data set is present from a previous selection.  Start the next set one past the end of the former set(s)
    io = .Cells(Rows.Count, 1).End(xlUp).Row + 1
    jo = .Cells(Rows.Count, 2).End(xlUp).Row + 1
    ko = .Cells(Rows.Count, 3).End(xlUp).Row + 1
End If

.Cells(io, 1).Value = "Date"
.Cells(jo, 2).Value = y1Label
.Cells(ko, 3).Value = y2Label
'goes through and prints the data that needs to be plotted to the worksheet tab
i = io + 1
For Each val In X
    .Cells(i, 1).Value = val
    i = i + 1
Next val

j = jo + 1

For Each val In y1
    .Cells(j, 2).Value = val
    j = j + 1
Next val

k = ko + 1

For Each val In y2
    .Cells(k, 3).Value = val
    k = k + 1
Next val

'If i = 1 Then GoTo ErrLine

FinalRow = .Cells(Rows.Count, 1).End(xlUp).Row ' record the end of the data

Set rng1 = Range(.Cells(1, 1), .Cells(FinalRow, 3)) 'set the range that contains the relevant data

rng1.Select 'excel is weird and requires the range to be selected in order to plot stugg
'create a plot object and set the plot details
'one series for cost, the other for labor hours or crew size or crew hours etc.
Set CH = .Shapes.AddChart2(Style:=201, _
XlChartType:=xlXYScatterSmooth, _
Left:=[E6].Left, _
Top:=[E6].Top, _
NewLayout:=True).Chart
'Be careful about the top and left attributes
Set Ser1 = CH.FullSeriesCollection(1)
Ser1.Name = y1Label

'Move Series 2 to secondary axis as smooth scatter plot
Set Ser2 = CH.FullSeriesCollection(2)
With Ser2
.Name = y2Label
.AxisGroup = xlSecondary
.ChartType = xlXYScatterSmooth
End With
'Format the plot
CH.SetElement msoElementLegendRight
CH.ChartTitle.Caption = ChartTitle
CH.Axes(xlCategory, xlPrimary).HasTitle = False 'Change to true for axis label
'CH.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Date"
'CH.Axes(xlCategory).AxisTitle.Font.Size = 14
CH.Axes(xlValue, xlPrimary).HasTitle = True
CH.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = y1Label
CH.Axes(xlValue, xlSecondary).HasTitle = True
CH.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = y2Label
CH.ChartArea.Height = CH.ChartArea.Height * 1.5
CH.ChartArea.Width = CH.ChartArea.Width * 1.5
CH.Axes(xlCategory).TickLabels.Orientation = 90

CH.PlotArea.Height = CH.PlotArea.Height * 0.8
'plot the axis as day and change the major scale units
If Day = True Then
    If Ser1.Points.Count > 1 Then
        CH.Axes(xlCategory).MajorUnit = 7
        CH.Axes(xlCategory).MinorUnit = 1
    End If
End If

End With

Exit Function

'ErrLine:
'    MsgBox "No Graph to Display, Selection contains no data points "

End Function
