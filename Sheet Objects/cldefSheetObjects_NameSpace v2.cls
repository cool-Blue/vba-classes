VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cldefSheetObjects_NameSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const maxCells As Long = 1000
Private mWks As Worksheet
Private mSpaceName As String
Private mSpaceType As String
Private mVariables As Collection
Private mSpaceRange As Range
Private mHeaderRange As Range
Private mObjects As Collection
Public Function addVariableUnTyped(varName As String) As Boolean
Dim v As clSheetObjects_Variable

    If mObjects Is Nothing Then
        Set mObjects = New Collection
    Else
        If isIncluded(mObjects, varName) Then
            addVariableUnTyped = False
            Exit Function
        End If
    End If
    Set v = New clSheetObjects_Variant
    With v
        .init Me, varName
    End With
    mObjects.Add v, varName
    
End Function
Public Function addVariableTyped(varName As String, varType As String) As Boolean
Dim v As clSheetObjects_Variable

    If mObjects Is Nothing Then
        Set mObjects = New Collection
    Else
        If isIncluded(mObjects, varName) Then
            addVariableTyped = False
            Exit Function
        End If
    End If
    Set v = New clSheetObjects_Variable
    With v
        .init Me, varName
        If UCase(Trim(varType)) <> UCase(Trim(.varType)) Then
            addVariableTyped = False
            Exit Function
        End If
    End With
    mObjects.Add v, varName
    
End Function
Property Get Variable(varName As String) As clSheetObjects_Variable
Attribute Variable.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Variable.VB_UserMemId = 0
    Set Variable = mObjects.Item(varName)
End Property
Property Let nameSpace(wks As Worksheet, spaceType As String, Name As String)
    Set mWks = wks
    mSpaceType = spaceType
    mSpaceName = Name
End Property
Property Get Name() As String
    Name = mSpaceName
End Property
Property Get Header() As Range
    If mHeaderRange Is Nothing Then _
        Set mHeaderRange = getHeaderRange(mWks, mSpaceType, mSpaceName)
    Set Header = mHeaderRange
End Property
Property Get SpaceRange() As Range
    If mSpaceRange Is Nothing Then _
        Set mSpaceRange = getSpaceRange(mWks, mSpaceType, mSpaceName)
    Set SpaceRange = mSpaceRange
End Property
Private Function getSpaceRange(wks As Worksheet, spaceType As String, Name As String) As Range
    Set getSpaceRange = Header
    Set getSpaceRange = getSpaceRange.Resize(Rowsize:=currentRegionBorders(getSpaceRange.Cells(1), xlDown).Rows.Count)
End Function
Private Function getHeaderRange(wks As Worksheet, spaceType As String, Name As String) As Range
Dim pTitle As Range
    Set pTitle = wks.Cells.Find(Name)
    Do While Trim(pTitle.Offset(0, 1)) <> spaceType
        Set pTitle = wks.Cells.FindNext
    Loop
    Set getHeaderRange = currentRegionBorders(pTitle, xlToRight)
End Function

Private Function currentRegionBorders(startCell As Range, Optional direction As XlDirection = xlToLeft, Optional limit As Long = maxCells) As Range
Dim i As Long, c As Range, soughtEdge As XlBordersIndex, searchEdge As XlBordersIndex

    searchEdges direction, soughtEdge, searchEdge
    Set c = startCell
    For i = 0 To limit
        nextCell c, direction
        With startCell.Borders(soughtEdge)
            If c.Borders(searchEdge).LineStyle = .LineStyle Then
                If c.Borders(searchEdge).Weight = .Weight _
                    And c.Borders(searchEdge).Color = .Color Then Exit For
            End If
        End With '.Borders(xlEdgeRight)
    Next i
    Set currentRegionBorders = mWks.Range(startCell, c)
End Function
Private Sub nextCell(ByRef c As Range, Optional direction As XlDirection = xlToLeft)
Dim maxRow As Long, maxCol As Long

    If mWks Is Nothing Then Exit Sub
    maxRow = mWks.Rows.CountLarge: maxCol = mWks.Columns.CountLarge
    
    Select Case direction
    Case xlDown
        If c.Row < maxRow Then Set c = c.Offset(1, 0)
    Case xlToLeft
        If c.Column > 1 Then Set c = c.Offset(0, -1)
    Case xlToRight
        If c.Column < maxCol Then Set c = c.Offset(0, 1)
    Case xlUp
         If c.Row > 1 Then Set c = c.Offset(-1, 0)
    End Select
    
End Sub
Private Sub searchEdges(direction As XlDirection, _
                        soughtEdge As XlBordersIndex, _
                        searchEdge As XlBordersIndex)
    Select Case direction
    Case xlDown
        soughtEdge = xlEdgeTop
        searchEdge = xlEdgeBottom
    Case xlToLeft
        soughtEdge = xlEdgeRight
        searchEdge = xlEdgeLeft
    Case xlToRight
        soughtEdge = xlEdgeLeft
        searchEdge = xlEdgeRight
    Case xlUp
        soughtEdge = xlEdgeBottom
        searchEdge = xlEdgeTop
    End Select
End Sub
Property Get varNameRange(varName As String) As Range
    Set varNameRange = mSpaceRange.Find(varName)
End Property
Private Function isIncluded(col As Variant, key As String) As Boolean
Const NOT_INCLUDED As Long = 5
    On Error GoTo incol
    col.Item key
incol:
    isIncluded = (Err.Number <> NOT_INCLUDED)
End Function

Private Sub Class_Terminate()
    On Error Resume Next     'allow for Collection is already Nothing
    Set mObjects = Nothing
End Sub
