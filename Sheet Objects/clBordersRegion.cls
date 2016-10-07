VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clBordersRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const maxCells As Long = 1000
Public Function currentRegionBorders(startCell As Range, _
                                        Optional direction As XlDirection = xlToLeft, _
                                        Optional limit As Long = maxCells, _
                                        Optional guideEdge As XlBordersIndex = xlEdgeTop) As Range
                                        
Dim i As Long, c As Range, soughtEdge As XlBordersIndex, searchEdge As XlBordersIndex, wks As Worksheet

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
    Set wks = startCell.Worksheet
    Set currentRegionBorders = wks.Range(startCell, c)
End Function
Private Sub nextCell(ByRef c As Range, Optional direction As XlDirection = xlToLeft)
Dim maxRow As Long, maxCol As Long, wks As Worksheet

    Set wks = c.Worksheet
    maxRow = wks.Rows.CountLarge: maxCol = wks.Columns.CountLarge
    
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


