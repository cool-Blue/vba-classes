VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBlockDiagramBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const labelTag As String = "_BlockDiagram_Label"

Private Enum bdElementType
  [_LBound] = 0
  bdChildren = bdElementType.[_LBound]
  bdHeads = 1
  bdTails = 2
  [_UBound] = bdElementType.bdTails
End Enum

'Sheet surface structure
Private mShape As Shape
Private mCell As Range

'User interaction structure
Private mOLE As OLEObject
Private WithEvents mLb As MSForms.Label
Attribute mLb.VB_VarHelpID = -1

'Related and connected elements
Private mElements(bdElementType.[_UBound]) As cCollection

'Instance State
Private mHeadBlock As cBlockDiagramHeadBlock
Private mActivated As Boolean

'Messaging
Private mNotify As cNotify
Public Property Let Notifier(RHS As cNotify)
  Set mNotify = RHS
End Property

'Related and connected elements management
Property Let addHead(RHS As String)
Const myName As String = "cBlockDiagramBlock.addTail"
Dim db As New cDebugReporter
    db.Report Caller:=IIf(Name = vbNullString, myName, Name & "." & Split(myName, ".")(1))

  addElement bdHeads, RHS

End Property
Property Let addTail(RHS As String)
Const myName As String = "cBlockDiagramBlock.addTail"
Dim db As New cDebugReporter
    db.Report Caller:=IIf(Name = vbNullString, myName, Name & "." & Split(myName, ".")(1))

  addElement bdTails, RHS

End Property
Property Let addChild(RHS As String)
Const myName As String = "cBlockDiagramBlock.addChild"
Dim db As New cDebugReporter
    db.Report Caller:=IIf(Name = vbNullString, myName, Name & "." & Split(myName, ".")(1))

  addElement bdChildren, RHS

End Property
Private Sub addElement(elementType As bdElementType, element As String)
Const myName As String = "cBlockDiagramBlock.addLink"
Dim db As New cDebugReporter
    db.Report Caller:=IIf(Name = vbNullString, myName, Name & "." & Split(myName, ".")(1))

  If mElements(elementType) Is Nothing Then
    db.Report message:="First Element: " & element
    Set mElements(elementType) = New cCollection
  End If
  mElements(elementType).Add element, element
  
  db.exitMessage = mElements(elementType).Count
End Sub

'Instance identification
Property Get elementType() As cbElement
  elementType = cbBlock
End Property
Property Get Name() As String
  If hasParent Then
    Name = ParentName & " " & Cell.Value
  Else
    If Not Cell Is Nothing Then
      Name = Cell.Value
    Else
      Name = vbNullString
    End If
  End If
End Property
Public Property Get ParentName() As String
  If hasParent Then
    ParentName = ParentCell.Value
  Else
    ParentName = vbNullString
  End If
End Property
Private Property Get ParentCell() As Range
  If Not Cell Is Nothing Then Set ParentCell = Cell.CurrentRegion.Cells(1)
End Property
Public Property Get hasParent() As Boolean
  If Not (ParentCell Is Nothing Or Cell Is Nothing) Then
    hasParent = Application.Intersect(ParentCell, Cell) Is Nothing
  End If
End Property

'Sheet surface structure accessors
Public Property Let Shape(RHS As Shape)
  Set mShape = RHS
End Property
Property Get Shape() As Shape
  Set Shape = mShape
End Property
Property Get Cell() As Range
  If Not mShape Is Nothing Then Set mCell = mShape.TopLeftCell
  Set Cell = mCell
End Property
Private Property Get ws() As Worksheet
  Set ws = Cell.Worksheet
End Property

Property Let element(r As Shape)
Const myName As String = "cBlockDiagramBlock.Rectangle"
Dim db As New cDebugReporter
    db.Report Caller:=IIf(Name = vbNullString, myName, Name & "." & Split(myName, ".")(1))

  Set mShape = r
  applyLabel
  db.exitMessage = Name
  
End Property
Private Sub applyLabel()
Const myName As String = "cBlockDiagramBlock.applyLabel"
Dim db As New cDebugReporter
    db.Report Caller:=IIf(Name = vbNullString, myName, Name & "." & Split(myName, ".")(1))

Dim n As String

  If existingLabel Is Nothing Then
    Set mOLE = ws.OLEObjects.Add(classtype:="Forms.Label.1", _
                                  Top:=Cell.Top, Left:=Cell.Left, _
                                  Height:=Cell.Height, Width:=Cell.Width)
    
    n = HashIt(Me.Name & " " & labelTag)
    mOLE.Name = n
    If mOLE.ShapeRange(1).Name <> n Then mOLE.ShapeRange(1).Name = n
  Else
    Set mOLE = existingLabel
  End If
  
  
  Set mLb = mOLE.Object
  With mLb
    .Caption = ""
    .BackStyle = fmBackStyleTransparent
    .Interior.ColorIndex = xlColorIndexNone
  End With
  
  db.exitMessage = mLb.Name
End Sub
Private Property Get existingLabel() As OLEObject
Dim b As Boolean, o As OLEObject

  On Error GoTo exitFalse
  Set o = ws.OLEObjects(HashIt(Me.Name & " " & labelTag))
  On Error GoTo 0
  
  b = o.Top = Cell.Top
  b = b And o.Left = Cell.Left
  b = b And o.Height = Cell.Height
  b = b And o.Width = Cell.Width
  If b Then Set existingLabel = o
  Exit Property
exitFalse:
  Err.Clear
End Property

'State management and comunications
Public Property Let headBlock(Backwards As Boolean, RHS As Boolean)
Const myName As String = "cBlockDiagramBlock.headBlock"
Dim db As New cDebugReporter
    db.Report Caller:=IIf(Name = vbNullString, myName, Name & "." & Split(myName, ".")(1))
  
' Start propogation from this instance in the direction indicated.
' Switch on or switch off depending on RHS.
' The HeadBlock object provides virtual sender (since there is no sender for a head block).
  mHeadBlock = RHS
  
  If RHS Then
    Activate mHeadBlock.Sender(Backwards)
  Else
    deActivate mHeadBlock.LastSender
  End If
  
End Property
Private Property Get propogateBackwards(Sender As String) As Boolean
' Infer the direction of propogation by checking the sender object

' Check for pier senders
  If Not mElements(bdTails) Is Nothing Then
    propogateBackwards = mElements(bdTails).isIncluded(Sender)
  End If
  
' Special signaling for Head Block (triggered by user input directly)
  propogateBackwards = propogateBackwards Or mHeadBlock.Backwards(Sender)
  
End Property
Public Property Let Activated(RHS As Boolean)
  mActivated = RHS
End Property
Public Property Get Activated() As Boolean
  Activated = mActivated
End Property
Public Sub Activate(Sender As String)
Const myName As String = "cBlockDiagramBlock.Activate"
Dim db As New cDebugReporter
    db.Report Caller:=IIf(Name = vbNullString, myName, Name & "." & Split(myName, ".")(1)), _
              newLine:=Both
  
  Highlight
  notifyParent True
  mNotify.SendMessage "notifyElements", Me.Name, mElements(bdChildren), True
  
  If propogateBackwards(Sender) Then
    mNotify.SendMessage "notifyElements", Me.Name, mElements(bdHeads), True
  Else
    mNotify.SendMessage "notifyElements", Me.Name, mElements(bdTails), True
  End If
  
End Sub

Public Sub deActivate(Sender As String)
Const myName As String = "cBlockDiagramBlock.deActivate"
Dim db As New cDebugReporter
    db.Report Caller:=IIf(Name = vbNullString, myName, Name & "." & Split(myName, ".")(1))
    
  unHighlight
  notifyParent False
  mNotify.SendMessage "notifyElements", Me.Name, mElements(bdChildren), False
  
  If propogateBackwards(Sender) Then
    mNotify.SendMessage "notifyElements", Me.Name, mElements(bdHeads), False
  Else
    mNotify.SendMessage "notifyElements", Me.Name, mElements(bdTails), False
  End If
  
End Sub
Public Property Let Highlighted(RHS As Boolean)
  If RHS Then
    If hasParent Then
      Highlight
    Else
      Highlight
    End If
  Else
    unHighlight
  End If
End Property
Private Sub Highlight()

  With mShape.line
    .Visible = msoTrue
    .ForeColor.Brightness = 0
    .Transparency = 0
    If hasParent Then
      .ForeColor.ObjectThemeColor = msoThemeColorAccent2
      .ForeColor.tintAndShade = 0
      .weight = 1
    Else
      .ForeColor.ObjectThemeColor = msoThemeColorBackground1
      .ForeColor.tintAndShade = -0.25
      .weight = 2
    End If
  End With
  
End Sub
Private Sub unHighlight()
  mShape.line.Visible = msoFalse
End Sub

Public Function isQualifiedLabel(o As OLEObject) As Boolean
Dim s As String
  s = o.Name
  isQualifiedLabel = (InStr(s, labelTag) > 0)
End Function

Private Sub notifyParent(Activate As Boolean)
  If hasParent Then mNotify.SendMessage "notifyParent", vbNullString, ParentName, Activate 'mNotify.Parent ParentName, Activate
End Sub

'UI Events management
Private Sub mLb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  If Not mHeadBlock And Button <> xlSecondaryButton Then
    Dim mApp As New cApplication
    
    mApp.quietMode
    On Error GoTo exitEnabled
    mNotify.SendMessage "notifyActivated", Me.Name, vbNullString, True
  
  End If
exitEnabled:
End Sub

Private Sub mLb_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Const myName As String = "cBlockDiagramBlock.mLb_MouseUp"
Dim db As New cDebugReporter
    db.Report Caller:=IIf(Name = vbNullString, myName, Name & "." & Split(myName, ".")(1))
  Dim mApp As New cApplication

  Select Case Button
  
  Case xlSecondaryButton
    mApp.quietMode
    On Error GoTo exitEnabled
    mNotify.SendMessage "notifyActivated", Me.Name, vbNullString, False
    
  Case xlPrimaryButton
    mApp.quietMode
    On Error GoTo exitEnabled
    mNotify.SendMessage "notifyActivated", Me.Name, vbNullString, True
    
  End Select
  
exitEnabled:
End Sub

'Life cycle management
Private Sub Class_Initialize()
  Set mHeadBlock = New cBlockDiagramHeadBlock
End Sub

Private Sub Class_Terminate()
Const myName As String = "cBlockDiagramBlock.Class_Terminate"
Dim db As New cDebugReporter
    db.Report Caller:=myName


  On Error Resume Next
  If Not mOLE Is Nothing Then mOLE.Delete
End Sub

