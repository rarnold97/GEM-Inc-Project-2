'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Public Function SortDict(dctList As Object) As Object
    'Sorts Dictionary By Date
    'Retrieved From : https://stackoverflow.com/questions/14808104/sorting-a-dictionary-by-key-in-vba
    
    Dim arrTemp() As Date
    Dim curKey As Variant
    Dim itX As Integer
    Dim itY As Integer

    'Only sort if more than one item in the dict
    If dctList.Count > 1 Then

        'Populate the array
        ReDim arrTemp(dctList.Count)
        itX = 0
        For Each curKey In dctList
            arrTemp(itX) = curKey
            itX = itX + 1
        Next

        'Do the sort in the array
        For itX = 0 To (dctList.Count - 2) ' see the logic behind a bubble sort.  Goes to -2 to be able to index list entrys ahead
            For itY = (itX + 1) To (dctList.Count - 1)
                If arrTemp(itX) < arrTemp(itY) Then
                    curKey = arrTemp(itY)
                    arrTemp(itY) = arrTemp(itX)
                    arrTemp(itX) = curKey
                End If
            Next
        Next

        'Create the new dictionary
        Set SortDict = CreateObject("Scripting.Dictionary")
        For itX = 0 To (dctList.Count - 1)
            SortDict.Add arrTemp(itX), dctList(arrTemp(itX))
        Next

    Else
        Set SortDict = dctList 'return the dictionary to the main code
    End If
    
End Function
