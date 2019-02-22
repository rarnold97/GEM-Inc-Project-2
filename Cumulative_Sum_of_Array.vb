'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Public Function CumSum(col As Collection) As Collection
'scans through and array and literally takes the cumulative sum
'the whole reverse array thing is a hard code to fix the fact that excel reads this set backwards
'kind of confusing honestly
Dim i, j As Integer
Dim arr(1 To col.Count) As Variant
Dim NewCol As Collection

Set NewCol = New Collection

j = 2

arr(1) = col(col.Count)

For i = col.Count To 1 Step -1
    arr(j) = arr(j - 1) + col(i)
Next i

arr = ReverseArray(arr)
'steps through essentially generates a discrete cumulative distribution
For Each Item In arr
    NewCol.Add Item
Next Item

CumSum = NewCol

End Function
