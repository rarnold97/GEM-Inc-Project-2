'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Public Function Charts2PDF(PassedName As String)

    Dim Ws As Worksheet, wsTemp As Worksheet
    Dim tp As Long
    Dim NewFileName As String
    Dim chrt As ChartObject
    Dim myfile As Variant
    Dim shp As Shape
    Dim lCnt As Long
    Dim dTop As Double
    Dim dLeft As Double
    Dim dHeight As Double
    Dim Space As Long
    DateString = Format(Now, "yyyy-mm-dd")
    
    NewFileName = Application.ActiveWorkbook.Path & "\" & "PM_Report_Graphs" & " " & DateString & ".pdf"
    'The user form I created automatically activates the sheet selected by the user that contains the graphs

    Set Ws = Sheets(PassedName) 'set the worksheet to the one that contains the plots

    Application.ScreenUpdating = False 'basically turns off annoying stuff this is optional
    
    tp = 10 ' pre defines the graph position
    Space = 50 'number of space units between graphs to manipulate them in to printing all to one page
    
    Set wsTemp = Sheets.Add 'add a temporary blank sheet
    
    wsTemp.PageSetup.Orientation = xlLandscape 'change the orientation to landscape
    
    For Each shp In Ws.Shapes 'loop through all the charts in the sheet
            shp.Copy 'make a copy
            wsTemp.Range("A1").PasteSpecial 'paste to the new sheet
            Selection.Top = tp 'set the inital coordinates
            Selection.Left = 5
            
            If Selection.TopLeftCell.Row > 1 Then 'check to see if it is the first pasted chart
                wsTemp.Rows(Selection.TopLeftCell.Row).PageBreak = xlPageBreakManual 'set a page break to print to the next page.  Tricking excel
            End If
            tp = tp + Selection.Height + 50 ' add to the spacing between charts
    Next
    'save to a pdf
    myfile = Application.GetSaveAsFilename(InitialFileName:=NewFileName, FileFilter:="PDF Files(*.pdf),*.pdf", Title:="Select Folder and File Name to Save as PDF")
    'check if file wasnt closed or cancelled when the print dialog is opened
    If myfile <> False Then
        wsTemp.ExportAsFixedFormat Type:=xlTypePDF, Filename:=myfile, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Else
        MsgBox "No File Selected. PDF will not be saved", vbOKOnly, "No File Selected" 'inform user that they did not accept the print dialog box
    End If
    'turn off alerts
    Application.DisplayAlerts = False
    'delete the temporary sheet
    wsTemp.Delete
    'turn screen updating and alerts back on.  Essentially, without this, it will be pesky and say hey do you really want to delete this sheet?
    'and of course we do because it is a temporary sheet.  The idea is all this happens behind the scenes without the user seeing, hence
    'why we turn screen updaing and alers off temporarily and then turn them back on
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

End Function
