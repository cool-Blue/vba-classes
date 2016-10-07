Attribute VB_Name = "modSheetGroupsArray"
Option Explicit
Sub loadNameKeys(ByRef nameKeys() As Variant, ByRef sheetGroups() As Collection)
'Assume Selection is the first column of a jagged list of string name keys of arbritary length.
'No blanks are allowed in the lists
Dim keySet() As String, lKeySetCount As Long, k As Long, r As Long
Dim rGroup As Range, c As Range, col1 As Range, groupCount As Long

    Set col1 = Selection
'   Set the first dimension of the data objects based on the number of groups
    groupCount = col1.Rows.Count
    ReDim nameKeys(groupCount - 1)
    ReDim sheetGroups(groupCount - 1)
    
    r = 0
'   Read the spreadsheet table into an array of 1D string arrays
    For Each c In col1
'       Get the next row of common keys
        With c
            lKeySetCount = .End(xlToRight).Column - .Column + 1
            Set rGroup = .Resize(1, lKeySetCount)
        End With 'c
'       Load the key set into an array of strings
        ReDim keySet(lKeySetCount - 1)
        For k = 0 To lKeySetCount - 1
            keySet(k) = rGroup.Cells(1, k + 1)
        Next k
'       Add the key set to the array of key sets
        nameKeys(r) = keySet
        Set sheetGroups(r) = New Collection
        r = r + 1
    Next c

End Sub
Sub loadSheetGroups(wkb As Workbook, nameKeys() As Variant, ByRef psheetGroups() As Collection)
Dim sh As Worksheet, g As Long, k As Long
Dim groupCount As Long, nameKey As Variant, nameGroup() As String, lKeySetCount As Long
Dim groupName As String
    
    groupCount = UBound(nameKeys, 1)
    
    For Each sh In wkb.Sheets
        For g = 0 To groupCount
            groupName = Join(nameKeys(g), "_")
            lKeySetCount = UBound(nameKeys(g))
            For Each nameKey In nameKeys(g)
                If Left(sh.Name, 3) = nameKey Then
                    psheetGroups(g).Add sh, sh.CodeName
                End If
            Next nameKey
        Next g
    Next sh
    
End Sub
Sub test()
Dim sheetGroups() As Collection
Dim pnameKeys() As Variant, sg As Variant
Dim sh As Worksheet, group As Variant, g As Long
Dim message As String

    loadNameKeys pnameKeys, sheetGroups
    loadSheetGroups ActiveWorkbook, pnameKeys, sheetGroups
    g = 0
    For Each group In pnameKeys
        message = message & Join(group, "_") & vbCrLf
        For Each sh In sheetGroups(g)
            message = message & vbTab & sh.Name & vbCrLf
        Next sh
        g = g + 1
    Next group
    MsgBox message
End Sub
Sub testImediate()
Dim sheetGroups() As Collection
Dim pnameKeys() As Variant, sg As Variant
Dim sh As Worksheet, group As Variant, g As Long

    loadNameKeys pnameKeys, sheetGroups
    loadSheetGroups ActiveWorkbook, pnameKeys, sheetGroups
    g = 0
    For Each group In pnameKeys
        Debug.Print Join(group, "_")
        For Each sh In sheetGroups(g)
            Debug.Print vbTab & sh.Name
        Next sh
        g = g + 1
    Next group
End Sub
Public Function isIncluded(col As Collection, key As String) As Boolean
Const NOT_INCLUDED As Long = 5
    On Error GoTo incol
    col.Item key

incol:
    isIncluded = (Err.Number <> NOT_INCLUDED)
End Function
