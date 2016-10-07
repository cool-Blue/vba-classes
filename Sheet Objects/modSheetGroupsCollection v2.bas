Attribute VB_Name = "modSheetGroupsCollection"
Option Explicit
Const Err_No_Selection As Long = vbObjectError + 1
Const Err_No_Selection_Message As String = "Dude: the selection is a blank cell..."
'///////////////////////////////////////////////////////////////////////////////////////
'Collection having members that will be Collections of Worksheet Objects
'   ...a jagged array of Worksheets
'The Key to the Collection of Collections will be the unique group name.
'   gsheetGroups.Item(myGroupName) returns a Collection of member Worksheets.
'The Key to each of the Collection of Worksheets will be the Sheet CodeName.
'   gsheetGroups.Item(myGroupName).Item(myCodeName) returns a member Worksheet
Public gsheetGroups As Collection
'///////////////////////////////////////////////////////////////////////////////////////
'A collection of the group names.  Used for enumerating the other collections.
'These are constructed by Joining the memebrs of the group into a single string.
Public ggroupNames As Collection
'///////////////////////////////////////////////////////////////////////////////////////
'A collection of 1D string arrays.
'   ...a jagged array of strings, with each array of Stings being a group of name keys.
'   Key is the groupName
Public gnameKeys As Collection
'///////////////////////////////////////////////////////////////////////////////////////
Sub loadNameKeys(rangeNameKeys As Range, ByRef nameKeys As Collection, _
                    ByRef groupNames As Collection)
'///////////////////////////////////////////////////////////////////////////////////////
' Procedure :   loadNameKeys
' Author    :   User
' Date      :   1/07/2014
' Purpose   :   Read a list on name keys from rangeNameKeys and load them into a Collection
'               Create a unique name for each group and load them into a Collection.  Also
'               use the group name as the Key field for the first collection.
'               Assume Selection is the first column of a jagged list of
'               string name keys of arbritary length.
'               No blanks are allowed in the lists... sorry
'//////////////////////////////////////////////////////////////////////////////////////
Const myName As String = "loadNameKeys"
Dim keySet() As String, lKeySetCount As Long, k As Long
Dim c As Range, r As Long
Dim groupName As String

'   Check if the selection is blank and throw an error if it is
    If Selection.Cells(1, 1).Value = "" Then
        Err.Raise Err_No_Selection, myName, Err_No_Selection_Message
    End If
    
'   Make it the first column only if not already
    Set rangeNameKeys = rangeNameKeys.Resize(rangeNameKeys.Rows.Count, 1)
    
    r = 0
'   Read the spreadsheet table into an array of 1D string arrays
    For Each c In rangeNameKeys
'       extend c to include the current row of common keys
        With c
            lKeySetCount = .End(xlToRight).Column - .Column + 1
            Set c = .Resize(1, lKeySetCount)
    '       Load the key set into an array of strings
    '       Prepare the array...
            ReDim keySet(lKeySetCount - 1)
    '       and load it
            For k = 0 To lKeySetCount - 1
                keySet(k) = .Cells(1, k + 1)
            Next k
        End With 'c
'       Make a unique name for each group and use it as the Key for
'       the group names collection
        groupName = Join(keySet, "_")
        groupNames.Add groupName, groupName
'       Then load into the array of key sets into the collection
        nameKeys.Add keySet, groupName
    Next c

End Sub
Sub loadSheetGroups(wkb As Workbook, nameKeys As Collection, ByRef psheetGroups As Collection, ByRef pgroupNames As Collection)
'
Const ALREADY_EXISTS As Long = 5
Dim SheetGroup As Collection
Dim sh As Worksheet
Dim nameKey As Variant
Dim groupName As Variant
    
'   Check the name of every sheet in wbk...
    For Each sh In wkb.Sheets
'       and test for its membership of each group
        For Each groupName In pgroupNames
'           Test membersheet by comparing with all of the strings in the group
            For Each nameKey In nameKeys.Item(groupName)
                If Right(sh.Name, 3) = nameKey Then
'               If the sheet qualifies, add it to the .Item(groupName) collection.
'               If this is the first sheet to qualify for this group, then the collection
'               will not exist yet, so catch the error and add a new collection for this group.
                    On Error GoTo newGroup
                        psheetGroups.Item(groupName).Add sh, sh.CodeName
                    On Error GoTo 0
                End If
            Next nameKey
        Next groupName
    Next sh
    Exit Sub
newGroup:
'   If its not the duplicate Key error then pass the error to the caller unhandled.
    If Err.Number = ALREADY_EXISTS Then
        Set SheetGroup = New Collection
        psheetGroups.Add SheetGroup, groupName
        Resume
    End If
End Sub
Sub initCollections()
    Set gsheetGroups = New Collection
    Set ggroupNames = New Collection
    Set gnameKeys = New Collection
    
    loadNameKeys Selection, gnameKeys, ggroupNames
    loadSheetGroups ActiveWorkbook, gnameKeys, gsheetGroups, ggroupNames
End Sub

Sub test()
Dim sh As Worksheet, groupName As Variant
Dim message As String

'   Enumerate the membership of each group
    For Each groupName In ggroupNames
        message = message & groupName & vbCrLf
        For Each sh In gsheetGroups.Item(groupName)
            message = message & vbTab & sh.Name & vbCrLf
        Next sh
    Next groupName
    MsgBox message
    Debug.Print message
'   Enumerate the membership of each eligible sheet
    message = ""
    For Each sh In ActiveWorkbook.Sheets
        For Each groupName In ggroupNames
            If isIncluded(gsheetGroups.Item(groupName), sh.CodeName) Then
                message = message & sh.Name & " is in group " & groupName & vbCrLf
            End If
        Next groupName
    Next sh
    MsgBox message
    Debug.Print message
End Sub
Sub testParse()
Dim sh As Worksheet


    For Each sh In gsheetGroups.Item("ABC_123_jkjkl")
        sh.Columns("E:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        With sh.Columns("D:D")
                      .TextToColumns Destination:=sh.Range("D1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                            ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=True, _
                            OtherChar:="+", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
        End With
        sh.Columns("D:D").Delete Shift:=xlToLeft
        sh.Columns("F:F").Delete Shift:=xlToLeft
        sh.Columns("D:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        With sh.Columns("C:C")
                     .TextToColumns Destination:=sh.Range("C1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                           ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=True, _
                           OtherChar:="+", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        End With
        sh.Columns("E:F").Delete Shift:=xlToLeft
            sh.Columns("C:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        With sh.Columns("B:B")
                     .TextToColumns Destination:=sh.Range("B1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                           ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=True, _
                           OtherChar:=">", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        End With
    Next sh
End Sub
Public Function isIncluded(col As Variant, key As String) As Boolean
Const NOT_INCLUDED As Long = 5
    On Error GoTo incol
    col.Item key

incol:
    isIncluded = (Err.Number <> NOT_INCLUDED)
End Function
