'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Public Function ToDate(OrigDate As String) As Date
'This function is only necessary because CGC makes some data sets have a funky date format without dashes of slashes
Dim year, month, Day, YMD As String
' string extraction
year = Left(OrigDate, 4)

month = Mid(OrigDate, 5, 2)

Day = Right(OrigDate, 2)

YMD = month & "/" & Day & "/" & year
'convert appropriate entries to a date and return to the main code
ToDate = CDate(YMD)

End Function
