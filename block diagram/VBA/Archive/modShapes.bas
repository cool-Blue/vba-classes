Attribute VB_Name = "modShapes"
Option Explicit
Dim bd As cBlockDiagram

Private Sub testNodes()
Dim s As Shape
  Selection.BottomRightCell.Select
  Debug.Print s.Name
End Sub
Private Sub logShapes()
Dim s As Shape, opSheet As Worksheet, c As Range, tl As Range

  Application.Calculation = xlCalculationManual
  
  Set opSheet = ThisWorkbook.Sheets("Shapes")
  Set c = opSheet.Cells(1, 1)
  Set c = writeLine(c, "s.Name", "s.Top", "s.Left", "s.TopLeftCell.Value", "s.TopLeftCell.Top", "s.TopLeftCell.Left", "Head", "Tail")
  For Each s In ActiveSheet.Shapes
    Set tl = s.TopLeftCell
    Select Case s.AutoShapeType
    Case msoShapeRectangle
      Debug.Print s.Name, s.Top & " " & s.Left, tl.Value, tl.Top & " " & tl.Left
'      s.Top = tl.Top: s.Left = tl.Left: s.Width = tl.Width: s.Height = tl.Height
      Set c = writeLine(c, s.Name, s.Top, s.Left, tl.Value, tl.Top, tl.Left)
    Case msoShapeMixed
      Set c = writeLine(c, s.Name, "", "", "", "", "", s.ConnectorFormat.BeginConnectedShape.Name, s.ConnectorFormat.EndConnectedShape.Name)
    End Select
  Next s
  
  Application.Calculation = xlCalculationAutomatic
  
End Sub
Private Function writeLine(c As Range, ParamArray output() As Variant) As Range
Dim i As Long
  For i = 0 To UBound(output)
    c.Offset(0, i).Value2 = output(i)
  Next i
  Set writeLine = c.Offset(1)
End Function

Private Sub formatConnectors()
Const myName As String = "modShapes.formatConnectors"
Dim db As New cDebugReporter
    db.Report Caller:=myName

Dim s As Shape
  Application.ScreenUpdating = False
  For Each s In ActiveSheet.Shapes
    On Error GoTo errorHandler
    If s.ConnectorFormat.Type = msoConnectorCurve Then
      db.Report message:=s.ConnectorFormat.Type
      With s.line
        .BeginArrowheadStyle = msoArrowheadOval
        .Transparency = 0.3000000119
        .EndArrowheadLength = msoArrowheadLong
        .EndArrowheadWidth = msoArrowheadWide
        .Weight = 1.5
      End With
    End If
ignoreShape:
      db.Report message:="afterResume: " & Err.Description
  Next s
  Application.ScreenUpdating = True
  Exit Sub
errorHandler:
  If Err.Number = -2147024809 Then Resume ignoreShape
End Sub
Private Sub testBlockDiagram()
Dim bd As New cBlockDiagram
  Set bd = New cBlockDiagram
  bd.Ws = ThisWorkbook.Sheets("Shapes")
End Sub
Sub testShapes()
Dim s As Shape, index As Long
  For Each s In ActiveSheet.Shapes
    index = index + 1
    Select Case s.Type
    Case msoEmbeddedOLEObject
      Debug.Print index, s.Name & " msoEmbeddedOLEObject"
    Case msoOLEControlObject
      Debug.Print index, s.Name & " msoOLEControlObject"
    End Select
  Next s
End Sub

Public Sub testBD()
Attribute testBD.VB_ProcData.VB_Invoke_Func = "D\n14"
  Set bd = New cBlockDiagram
  bd.Ws = ThisWorkbook.Sheets("Stock Calc Chart")
End Sub

