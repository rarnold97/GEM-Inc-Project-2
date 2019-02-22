'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Function ReverseArray(arr As Variant) As Variant
'Function that takes an array as input and reverses the order of its contents. Retrieved From : https://stackoverflow.com/questions/40563940/vba-reverse-an-array
    Dim val As Variant

    With CreateObject("System.Collections.ArrayList") '<-- create a "temporary" array list with late binding
        For Each val In arr '<--| fill arraylist
            .Add val
        Next val
        .Reverse '<--| reverse it
        ReverseArray = .Toarray '<--| write it into an array
    End With
End Function
