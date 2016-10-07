VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBlockDiagramConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Enum bdcEndType
  bdHead
  bdTail
End Enum
Private mShape As Shape
Private mNotify As cNotify
Private mActivated As Boolean

Public Property Let Notifier(RHS As cNotify)
  Set mNotify = RHS
End Property

Public Property Get Name() As String
  Name = mShape.Name
End Property
Property Get elementType() As cbElement
  elementType = cbConnector
End Property

Property Get TailName() As String
  TailName = endName(bdTail)
End Property
Property Get HeadName() As String
  HeadName = endName(bdHead)
End Property
Private Function endName(connectorEnd As bdcEndType) As String
Dim fb As cBlockDiagramBlock

  If Not mShape Is Nothing Then
    Set fb = New cBlockDiagramBlock
    Select Case connectorEnd
    Case bdHead
      fb.Shape = mShape.ConnectorFormat.EndConnectedShape
    Case bdTail
      fb.Shape = mShape.ConnectorFormat.BeginConnectedShape
    End Select
    endName = fb.Name
  Else
    endName = vbNullString
  End If
  
End Function
Property Let element(RHS As Shape)
  Set mShape = RHS
End Property
Private Property Get propogateBackwards(sender As String) As Boolean
  propogateBackwards = sender = HeadName
End Property

Public Property Let Activated(RHS As Boolean)
  mActivated = RHS
End Property
Public Property Get Activated() As Boolean
  Activated = mActivated
End Property
Public Sub Activate(sender As String)
Const myName As String = "cBlockDiagramConnector.Activate"
Dim db As New cDebugReporter
    db.Report Caller:=myName, Context:=TailName & " to " & HeadName, newLine:=Both
  
  Highlight
  mNotify.elements True, IIf(propogateBackwards(sender), TailName, HeadName), Me.Name
  
End Sub
Public Sub deActivate(sender As String)
Const myName As String = "cBlockDiagramConnector.deActivate"
Dim db As New cDebugReporter
    db.Report Caller:=myName, Context:=HeadName

  mNotify.elements False, IIf(propogateBackwards(sender), TailName, HeadName), Me.Name
  unHighlight
  
End Sub
Private Sub Highlight()

  With mShape.line
    .weight = 3
    .Transparency = 0
  End With
  
End Sub
Public Sub unHighlight()
  With mShape.line
    .weight = 1.5
    .Transparency = 0.75
  End With
End Sub
