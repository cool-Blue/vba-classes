VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBlockDiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum cbElement
  cbBlock
  cbConnector
End Enum

Private mElements As cCollection
Private mMessageProcessor As cBlockDiagramMessageProcessor
Private WithEvents mNotifier As cNotify
Attribute mNotifier.VB_VarHelpID = -1
Private WithEvents mWs As Worksheet
Attribute mWs.VB_VarHelpID = -1

Public Property Let ws(RHS As Worksheet)
Const myName As String = "cBlockDiagram.Ws"
Dim db As New cDebugReporter
    db.Report Caller:=myName
Dim mApp As New cApplication

  Set mWs = RHS
  mApp.quietMode
  On Error GoTo exitEnabled
  cleanUp mWs
  loadElements mWs
  linkElements
exitEnabled:
End Property
Public Property Let preBuiltWs(RHS As Worksheet)
Const myName As String = "cBlockDiagram.Ws"
Dim db As New cDebugReporter
    db.Report Caller:=myName
Dim mApp As New cApplication

  Set mWs = RHS
  mApp.quietMode
  On Error GoTo exitEnabled
  loadElements mWs
  linkElements
exitEnabled:
End Property

Public Sub cleanUp(sourceSheet As Worksheet, Optional All As Boolean = False)
Const myName As String = "cBlockDiagram.cleanUp"
Dim db As New cDebugReporter
    db.Report Caller:=myName, newLine:=Before


Dim o As OLEObject, Os As OLEObjects, index As Long, tagDelete As Collection
Dim fb As cBlockDiagramBlock

  Set Os = sourceSheet.OLEObjects
  Set tagDelete = New Collection
  Set fb = New cBlockDiagramBlock
  
  For index = 1 To Os.Count
    Set o = Os.item(index)
    If fb.isQualifiedLabel(o) Or All Then
      db.Report message:="tagging " & o.Name
      tagDelete.Add o, CStr(index)
    End If
  Next index
  
  
  For Each o In tagDelete
    db.Report message:="type to delete " & TypeName(o)
    db.Report message:="deleting " & o.Name
    o.Delete
    DoEvents
  Next o
  
End Sub
Private Sub loadElements(sourceSheet As Worksheet)
Const myName As String = "cBlockDiagram.loadElements"
Dim db As New cDebugReporter
    db.Report Caller:=myName, newLine:=Before

Dim s As Shape, connectorType As Long
Dim con As cBlockDiagramConnector
Dim fb As cBlockDiagramBlock

  For Each s In sourceSheet.Shapes
    
    If s.Type = msoAutoShape Then
    
      connectorType = 0
      On Error GoTo errorHandler
      connectorType = s.ConnectorFormat.Type
      On Error GoTo 0
      
      db.Report message:=s.Name & " " & connectorType
      
      Select Case connectorType
      Case msoConnectorCurve, msoConnectorStraight
        db.Report message:="loading " & s.Name & " in mElements"
        Set con = New cBlockDiagramConnector
        con.element = s
        con.Notifier = mNotifier
        con.unHighlight
        mElements.Add con, con.Name
      Case Else
      
        Select Case s.AutoShapeType
        Case msoShapeRectangle
          Set fb = New cBlockDiagramBlock
          fb.element = s
          fb.Notifier = mNotifier
          fb.Highlighted = False
          mElements.Add fb, fb.Name
          db.Report message:=fb.Name & " loaded in mElements"
        Case Else
        End Select
        
      End Select
    End If
  Next s
  
  Exit Sub
  
errorHandler:
  Resume Next
End Sub
Private Sub linkElements()
Const myName As String = "cBlockDiagram.linkElements"
Dim db As New cDebugReporter
    db.Report Caller:=myName, message:="Elements :" & mElements.Count, newLine:=Before

Dim o As Object, c As cBlockDiagramConnector, b As cBlockDiagramBlock

  For Each o In mElements
    Select Case o.elementType
    Case cbConnector
    
      Set c = o
      db "Linking Connector " & o.Name & " Tail Block: "
            
      If Not c.TailName = vbNullString Then
        Set b = mElements(c.TailName)
        b.addTail = c.Name
      Else
        MsgBox o.Name & " tail is not connected", vbCritical
      End If
      
      If Not c.HeadName = vbNullString Then
        Set b = mElements(c.HeadName)
        b.addHead = c.Name
      Else
        MsgBox o.Name & " head is not connected", vbCritical
      End If
            
    Case cbBlock
    
      Set b = o
      If b.hasParent Then mElements(b.ParentName).addChild = b.Name
      
    Case Else
    End Select
  Next o
End Sub
Public Sub logShapes(opSheet As Worksheet)
Dim s As Shape, c As Range, tl As Range
Dim fb As cBlockDiagramBlock, con As cBlockDiagramConnector
Dim mApp As New cApplication

  mApp.quietMode
  
  Set c = opSheet.Cells(1, 1)
  Set c = writeLine(c, "s.Name", "s.Top", "s.Left", "Name", "s.TopLeftCell.Top", "s.TopLeftCell.Left", "Head", "Tail")
  Set fb = New cBlockDiagramBlock
  Set con = New cBlockDiagramConnector
  
  For Each s In ActiveSheet.Shapes
    If s.Type = msoAutoShape Then
      Set tl = s.TopLeftCell
      Select Case s.AutoShapeType
      Case msoShapeRectangle
        fb.Shape = s
  '      s.Top = tl.Top: s.Left = tl.Left: s.Width = tl.Width: s.Height = tl.Height
        Set c = writeLine(c, s.Name, s.Top, s.Left, fb.Name, tl.Top, tl.Left)
      Case msoShapeMixed
        con.element = s
        Set c = writeLine(c, s.Name, "", "", "", "", "", con.HeadName, con.TailName)
      End Select
    End If
  Next s
  
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
Dim mApp As New cApplication

  mApp.quietMode
  
  For Each s In ActiveSheet.Shapes
    On Error GoTo errorHandler
    If s.ConnectorFormat.Type = msoConnectorCurve Then
      db.Report message:=s.ConnectorFormat.Type
      With s.line
        .BeginArrowheadStyle = msoArrowheadOval
        .Transparency = 0.3000000119
        .EndArrowheadLength = msoArrowheadLong
        .EndArrowheadWidth = msoArrowheadWide
        .weight = 1.5
      End With
    End If
ignoreShape:
    db.Report message:="afterResume: " & Err.Description
  Next s
  
  Exit Sub
errorHandler:
  If Err.Number = -2147024809 Then Resume ignoreShape
End Sub


Private Sub mNotifier_Message(MessageType As String, Sender As String, Receiver As Variant, _
                              Action As Variant)
  mMessageProcessor.Process MessageType, Sender, Receiver, Action, mElements
End Sub

Private Sub mWs_SelectionChange(ByVal Target As Range)
Const myName As String = "cBlockDiagram.mWs_SelectionChange"
Dim db As New cDebugReporter
    db.Report Caller:=myName

  mNotifier.SendMessage "notifyActivated", vbNullString, vbNullString, False

End Sub
Private Function emptyString(s As String) As String
  emptyString = IIf(s = vbNullString, "nothing", s)
End Function

Public Property Get Count() As Long
  Count = mElements.Count
End Property
Private Sub Class_Initialize()
  Set mElements = New cCollection
  Set mNotifier = New cNotify
  Set mMessageProcessor = New cBlockDiagramMessageProcessor
End Sub

