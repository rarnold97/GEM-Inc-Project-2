'DISCLAIMER: This code is the propery of GEM Inc. and should only be modified/kept with company permission
'Author: Ryan Arnold, Project Engineer Co-Op
'Supervisor: John Marshall, Project Manager
'Date Completed 2/04/2019

Public Function MakeUnique(ByVal OrigArray As Variant) As Variant

'This Function is credited to Bill Jelen and Tracy Syrstad, Authors of Excel 2016 VBA and Macros
'Takes in an array of values, and returns only the unique values
'For the purposes of this program, this function will give unique identifiers for the welder tag, joint type, package name,
'and package spec

Dim vAns() As Variant ' Variant type is used when we are not initially sure of the data type
Dim lStartPoint As Long
Dim lEndPoint As Long
Dim lCtr As Long, lCount As Long
Dim iCtr As Integer
Dim col As New Collection
Dim sIndex As String
Dim vTest As Variant, vItem As Variant
Dim iBadVarTypes(4) As Integer
'Function does not work if array element is one of the following types
iBadVarTypes(0) = vbObject
iBadVarTypes(1) = vbError
iBadVarTypes(2) = vbDataObject
iBadVarTypes(3) = vbUserDefinedType
iBadVarTypes(4) = vbArray
'Check to see whether the parameter is an array
If Not IsArray(OrigArray) Then
    Err.Raise ERR_BP_NUMBER, , ERR_BAD_PARAMETER
    Exit Function
End If
lStartPoint = LBound(OrigArray) 'record the starting index of the array
lEndPoint = UBound(OrigArray) ' record the ending index of the array

For lCtr = lStartPoint To lEndPoint
    vItem = OrigArray(lCtr)
    'first check to see whether variable type is acceptable
    For iCtr = 0 To UBound(iBadVarTypes)
        If VarType(vItem) = iBadVarTypes(iCtr) Or VarType(vItem) = iBadVarTypes(iCtr) + vbVariant Then
            Err.Raise ERR_BT_NUMBER, , ERR_BAD_TYPE ' unacceptable inpuput, raise an error
            Exit Function
        End If
    Next iCtr
    'Add element to a collection, using it as the index
    'if an error occurs, the element already exists
    sIndex = CStr(vItem)
    'first element add automatically
    If lCtr = lStartPoint Then
        col.Add vItem, sIndex
        ReDim vAns(lStartPoint To lStartPoint) As Variant
        vAns(lStartPoint) = vItem
    Else
        On Error Resume Next
        'Elegant approach. If a duplicate value is detected the collection flags and then this skips over it
        'Value is then assigned to a regular array taking advantage of the nonduplicate nature of a collection
        col.Add vItem, sIndex
        If Err.Number = 0 Then
            lCount = UBound(vAns) + 1
            ReDim Preserve vAns(lStartPoint To lCount)
            vAns(lCount) = vItem
        End If
    End If
    Err.clear
Next lCtr ' go to the next value in the array
MakeUnique = vAns ' record the unique values


End Function


