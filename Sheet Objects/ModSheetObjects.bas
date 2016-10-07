Attribute VB_Name = "ModSheetObjects"
Option Explicit
Const Err_No_Selection As Long = vbObjectError + 1
Const Err_No_Selection_Message As String = "Dude: the selection is a blank cell..."

Public Type nameSpaceType
    spaceName As String
    spaceType As String
End Type

Public nameSpaceCollection As SheetObjects_Environment
Public nameSpaces As Collection
Dim nameSpaceNames As Collection
Dim variableNames As Collection

Dim globalNameSpace As SheetObjects_NameSpace
Dim varObj As SheetObjects_Variable
'///////////////////////////////////////////////////////////////////////////////////////
Sub loadNameSpaces(wks As Worksheet, groupNames As Range, ByRef Groups As Collection, ByRef Members As Collection)
'///////////////////////////////////////////////////////////////////////////////////////
' Procedure :   loadNameSpaces
' Author    :   User
' Date      :   1/07/2014
' Purpose   :   Read jagged array of strings to the right of groupNames into the Members Collection
'               Load groupNames into a Collection to be used as an enumerator.
'               Group Types are in groupNames.Offset(0,-1), Load These
'               Assume groupNames is the first column of a jagged list of
'               string name keys of arbritary length.
'               No blanks are allowed in the lists... sorry
'//////////////////////////////////////////////////////////////////////////////////////
Const myName As String = "loadNameSpaces"
Dim keySet() As String, memberCount As Long, k As Long
Dim c As Range, r As Range, varRow As Range
Dim mNS As nameSpaceType
Dim NS As SheetObjects_NameSpace

'   Check if the selection is blank and throw an error if it is
    If Selection.Cells(1, 1).value = "" Then
        Err.Raise Err_No_Selection, myName, Err_No_Selection_Message
    End If
    Set Groups = New Collection
    Set Members = New Collection
'   Make it the first column only if not already
    Set groupNames = groupNames.Resize(groupNames.Rows.Count, 1)
    
'   Read the spreadsheet table into an array of 1D string arrays
    For Each c In groupNames
        Set NS = New SheetObjects_NameSpace
        With c
            mNS.spaceName = .Text
            mNS.spaceType = .Offset(0, -1).Text
            NS.nameSpace wks, mNS.spaceType, mNS.spaceName
'           shift to the start of the members list
'           extend c to include the current row of common keys
            memberCount = .End(xlToRight).Column - .Column
            Set varRow = .Offset(0, 1).Resize(1, memberCount)
    '       Ceate a new variable in the Name Space
            For Each r In varRow
                NS.addVariableUnTyped r.Text
            Next r
        End With 'c
        Groups.Add NS.Name, NS.Name
        Members.Add NS, NS.Name
    Next c

End Sub
'///////////////////////////////////////////////////////////////////////////////////////
Sub loadNameSpaces2(wks As Worksheet, groupNames As Range, ByRef Members As SheetObjects_Environment)
'///////////////////////////////////////////////////////////////////////////////////////
' Procedure :   loadNameSpaces
' Author    :   User
' Date      :   1/07/2014
' Purpose   :   Read jagged array of strings to the right of groupNames into the Members Collection
'               Assume groupNames is the first column of a jagged list of
'               string name keys of arbritary length.
'               No blanks are allowed in the lists... sorry
'//////////////////////////////////////////////////////////////////////////////////////
Const myName As String = "loadNameSpaces"
Dim keySet() As String, memberCount As Long, k As Long
Dim c As Range, r As Range, varRow As Range
Dim mNS As nameSpaceType
Dim NS As SheetObjects_NameSpace

'   Check if the selection is blank and throw an error if it is
    If Selection.Cells(1, 1).value = "" Then
        Err.Raise Err_No_Selection, myName, Err_No_Selection_Message
    End If
    
'   Make it the first column only if not already
    Set groupNames = groupNames.Resize(groupNames.Rows.Count, 1)
    
'   Read the spreadsheet table into an array of 1D string arrays
    For Each c In groupNames
        Set NS = New SheetObjects_NameSpace
        With c
            mNS.spaceName = .Text
            mNS.spaceType = .Offset(0, -1).Text
            NS.nameSpace wks, mNS.spaceType, mNS.spaceName
'           shift to the start of the members list
'           extend c to include the current row of common keys
            memberCount = .End(xlToRight).Column - .Column
            Set varRow = .Offset(0, 1).Resize(1, memberCount)
    '       Ceate a new variable in the Name Space
            For Each r In varRow
                NS.addVariableUnTyped r.Text
            Next r
        End With 'c
        Members.Add NS
    Next c

End Sub

Sub test()
Dim rOutput As Range, spaceName As Variant, r As Range

    
    Set globalNameSpace = New SheetObjects_NameSpace
    With globalNameSpace
        .nameSpace ActiveSheet, "Module", "Global"
        .Header.Select
        .SpaceRange.Select
    End With 'globalNameSpace
    With globalNameSpace("no name")
        .nameRange.Select
        .Range.Select
        MsgBox .baseAddress
        Set rOutput = Application.InputBox("where do you want this stuff?", "Test", ActiveSheet.Cells(1, 7), Type:=8)
        Set rOutput = rOutput.Resize(.WordCount)
        rOutput = Application.Transpose(.Contents)
    End With
    
End Sub
Sub testGlobalVar()
Dim rOutput As Range, spaceName As Variant, r As Range
Dim nsgv As SheetObjects_GlobalRef
    
    Set nsgv = New SheetObjects_GlobalRef
    With nsgv.globalVar(ActiveSheet, "Global", "Module", "no name")
        .nameRange.Select
        .Range.Select
        MsgBox .baseAddress
        Set rOutput = Application.InputBox("where do you want this stuff?", "Test", ActiveSheet.Cells(1, 7), Type:=8)
        Set rOutput = rOutput.Resize(.WordCount)
        rOutput = Application.Transpose(.Contents)
    End With
    
End Sub

Sub testColl()
Dim nameSpace As SheetObjects_NameSpace
Dim rOutput As Range, spaceNames As Variant, s As Variant
    
    loadNameSpaces ActiveSheet, Selection, nameSpaceNames, nameSpaces
    For Each s In nameSpaceNames
        With nameSpaces.Item(s)
            .Header.Select
            .SpaceRange.Select
        End With 'nameSpace
    Next s

End Sub
Sub testColl2()
Dim nameSpace As SheetObjects_NameSpace
Dim rOutput As Range, spaceNames As Variant, s As Variant
Dim NS As SheetObjects_NameSpace, NSvar As SheetObjects_Variable
    
    Set nameSpaceCollection = New SheetObjects_Environment
    loadNameSpaces2 ActiveSheet, Selection, nameSpaceCollection
    For Each NS In nameSpaceCollection
        With NS
            .SpaceRange.Select
            If MsgBox(.Name & " Body", vbOKCancel) = vbCancel Then Exit For
            For Each NSvar In NS
                NSvar.Range.Select
                If MsgBox(NSvar.Name & " Body", vbOKCancel) = vbCancel Then Exit For
            Next NSvar
        End With 'NS
    Next NS

End Sub

