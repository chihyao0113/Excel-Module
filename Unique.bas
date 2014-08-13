Function Unique(DRange As Variant) As Variant
'Unique function that excluded the null value from the list
Dim dict As Object
Dim i As Long, j As Long, NumRows As Long, NumCols As Long
 
'Convert range to array and count rows and columns
If TypeName(DRange) = "Range" Then DRange = DRange.Value2
NumRows = UBound(DRange)
NumCols = UBound(DRange, 2)
 
'put unique data elements in a dictionay
Set dict = CreateObject("Scripting.Dictionary")
For i = 1 To NumCols
    For j = 1 To NumRows
        If IsEmpty(DRange(j, i)) Then
            'Drop the null value
        Else
            dict(DRange(j, i)) = 1
        End If
    Next j
Next i
 
'Dict.Keys() is a Variant array of the unique values in DRange
 'which can be written directly to the spreadsheet
 'but transpose to a column array first
 
Unique = WorksheetFunction.Transpose(dict.keys)
 
End Function

Function UniqueRow(DRange As Variant, Optional alignedAs As String = "column") As Variant
'Unique function that excluded the null value from the list
Dim dict As Object
Dim i As Long, j As Long, NumRows As Long, NumCols As Long
Dim values() As String

'Convert range to array and count rows and columns
If TypeName(DRange) = "Range" Then DRange = DRange.Value2
NumRows = UBound(DRange)
NumCols = UBound(DRange, 2)
 
'put unique data elements in a dictionay
Set dict = CreateObject("Scripting.Dictionary")
For i = 1 To NumRows
    ReDim values(1 To NumCols)
    For j = 1 To NumCols
        If j <= 1 Then
            joint = DRange(i, j)
        Else
            joint = joint & "|" & DRange(i, j)
        End If
        values(j) = DRange(i, j)
    Next j
    
    dict(joint) = values
Next i
 
For Each Item In dict.items
    For Each element In Item
        'Debug.Print element
    Next element
Next Item
 
'Dict.Keys() is a Variant array of the unique values in DRange
 'which can be written directly to the spreadsheet
 'but transpose to a column array first
 
'UniqueRow = WorksheetFunction.Transpose(Split(dict.keys, "|"))
'tmp = WorksheetFunction.Transpose(dict.items)

If alignedAs Like "row" Then
    UniqueRow = WorksheetFunction.Transpose(dict.items)
Else
    UniqueRow = dict.items
End If
End Function
