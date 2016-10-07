VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAnnimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Syntax: shapeName <ann property{-property} startValue to endValue { step s{ t}}{ rep r}
' \<([a-zA-Z_]+)\s+(\b[a-zA-Z-]+)\s+from\s+(\d+)\s+to\s+(\d+)\s+in\s+(\d*\.\d*|\d+)\s+rep\s+(\d+)
' $1: ann
' $2: properties delimited by '-'       : <[a-zA-Z_]+\s+(\b[a-zA-Z-]+)
' $3: start value of property
' $4: end value of property
' $5: cycle time
' $6: repeats
' Alternate aproach:
'   Properties: property{-nextproperty}*
'   "<[a-zA-Z_]+\s+(\b[a-zA-Z-]+)"
'   $1: properties delimited by '-'
'
'   Cycle descriptor: from d to d in t.t OR from d to d to d in t.t
'   "<[a-zA-Z_]+.*from\s+(\d+)\s+to\s+(\d+)\s+in\s+(\d*\.\d*|\d+)|from\s+(\d+)\s+to\s+(\d+)\s+to\s+(\d+)\s+in\s+(\d*\.\d*|\d+)"
'   $1: start value of property
'   $2: IF COUNT = 4 THEN Mid value of property ELSE EndValue
'   $3: IF COUNT = 4 THEN End Value of property ELSE cycle time
'   $4: IF COUNT = 4 THEN cycle time ELSE N/A
'
'   Repeats: rep d
'   "<[a-zA-Z_]+.*rep\s+(\d+)"
'   $1: integer repeats
'
'   Rewind: rew
'   "<[a-zA-Z_]+.*rew"
'   $1: "rew"
'   IF COUNT > 0 THEN animate with rewind
'
Const nameRegExp As String = "\<([a-zA-Z_]+)\s+(\b[a-zA-Z-]+)\s+from\s+(\d+)\s+to\s+(\d+)\s+in\s+(\d*\.\d*|\d+)\s+rep\s+(\d+)"
Const PROPERTIESregExp As String = "<[a-zA-Z_]+\s+(\b[a-zA-Z-]+)"
Const CYCLEregExpLong As String = "\<[a-zA-Z_]+.*from\s+(\d+)\s+to\s+(\d+)\s+to\s+(\d+)\s+in\s+(\d*\.\d*|\d+)"
Const CYCLEregExpShort As String = "\<[a-zA-Z_]+.*from\s+(\d+)\s+to\s+(\d+)\s+in\s+(\d*\.\d*|\d+)"
Const REPEATSregExp As String = "\<[a-zA-Z_]+.*rep\s+(\d+)"
Const REWINDregExp As String = "\<[a-zA-Z_]+.*rew"

Const defSTEP As Long = 20 'msec
Const defPERIOD As Long = 1000
Const defREVERSE As Boolean = False

Enum annAtributes
  annProperties = 1
  annStart = 2
  annEnd = 3
  annCycle = 4
  annRepeats = 5
End Enum

Private mShape As Shape
Private mShapeByName As cObjectByName
Private mProperties As String
Private mtPeriod As Long
Private mtStep As Long
Private mReverse As Boolean
Private mPhaseCount As Long
Private mRepeatCount As Long
Private mAttributes As SubMatches
Private WithEvents frameEngine As cAnnimateFrames
Attribute frameEngine.VB_VarHelpID = -1
Private mTextBox As Shape
Property Get Properties() As String
  Properties = getMatches(mShape.Name, PROPERTIESregExp)(0)
End Property
Property Get startValue() As Integer
  startValue = getMatches(mShape.Name, CYCLEregExpShort, CYCLEregExpLong)(0)
End Property
Property Get MidValue() As Integer
Dim SM As SubMatches
  Set SM = getMatches(mShape.Name, CYCLEregExpShort, CYCLEregExpLong)
  MidValue = IIf(longCycle(SM), SM(1), -1)
End Property
Property Get EndValue() As Integer
Dim SM As SubMatches
  Set SM = getMatches(mShape.Name, CYCLEregExpShort, CYCLEregExpLong)
  EndValue = SM(IIf(longCycle(SM), 2, 1))
End Property
Property Let tStep(RHS As Long)
  mtStep = tStep
End Property
Property Get tPeriod() As Long
Dim SM As SubMatches
  Set SM = getMatches(mShape.Name, CYCLEregExpShort, CYCLEregExpLong)
  On Error Resume Next
  tPeriod = SM(IIf(longCycle(SM), 3, 2)) * 1000
  On Error GoTo 0
  If tPeriod = 0 Then tPeriod = defPERIOD
End Property
Property Get tStep() As Long
  If mtStep = 0 Then tStep = defSTEP Else tStep = mtStep
End Property
Property Get Rewind() As Boolean
  Rewind = getMatches(mShape.Name, REWINDregExp)(0)
End Property
Property Get Repeats() As Integer
  On Error Resume Next
  Repeats = getMatches(mShape.Name, REPEATSregExp)(0)
  If Err <> 0 Then
    Err.Clear
    Repeats = 1
  End If
  On Error GoTo 0
End Property
Property Let Shape(RHS As Shape)
  Set mShape = RHS
  Set mShapeByName = New cObjectByName
  mShapeByName.baseObject = mShape
End Property
Property Get Shape() As Shape
  Set Shape = mShape
End Property
Property Let TextBox(RHS As Shape)
  Set mTextBox = RHS
End Property
Property Let ResolvedProperty(RHS As Variant)
  With mShapeByName
    .Properties = Properties
    .Value = RHS
  End With
End Property
Property Get ResolvedProperty() As Variant
  With mShapeByName
    .Properties = Properties
    ResolvedProperty = .Value
  End With
End Property
Sub Start(Optional resetRepeats As Boolean = True)
  mPhaseCount = 1
  If resetRepeats Then mRepeatCount = 1 Else mRepeatCount = mRepeatCount + 1
  runPhase startValue, IIf(MidValue <= 0, EndValue, MidValue)
End Sub

Sub runPhase(pStart As Integer, pEnd As Integer)
Const myName As String = "cAnnimate.Start"
Dim step As Integer, fstep As Double
Dim v As Integer, dt As Long
Dim dirn As Integer

  mShapeByName.Properties = Properties
    
  dirn = Sgn(pEnd - pStart)
  v = pStart: If v = 0 Then v = dirn '0 value breaks Glow object
  
  Set frameEngine = New cAnnimateFrames
  With frameEngine
    .TriggerText = mTextBox
    .ShapeByName = mShapeByName
    .startValue = pStart
    .EndValue = pEnd
    .StepDirection = dirn
    .dt = tStep
    .T = tPeriod
    .Start
  End With
    
End Sub

Private Sub Class_Initialize()
Const myName As String = "cAnnimate.Class_Initialize"
Dim db As New cDebugReporter
    db.Report caller:=myName


End Sub

Private Sub Class_Terminate()
Const myName As String = "cAnnimate.Class_Terminate"
Dim db As New cDebugReporter
    db.Report caller:=myName
  Set frameEngine = Nothing
End Sub

Private Sub frameEngine_annComplete()
Const myName As String = "cAnnimate.frameEngine_annComplete"
Dim db As New cDebugReporter
    db.Report caller:=myName, Message:=ResolvedProperty & vbTab & frameEngine.EndValue

  If MidValue > 0 Then
    If mPhaseCount = 1 Then
      mPhaseCount = 2
      runPhase MidValue, EndValue
      Exit Sub
    End If
  End If
  
  If mRepeatCount >= Repeats Then
    Set frameEngine = Nothing
  Else
    Start False
  End If
  
End Sub
Private Function getMatches(searchString As String, regExp As String, Optional regExp2 As String) As SubMatches
Dim RE As cRegExpression
  Set getMatches = Nothing
  Set RE = New cRegExpression
  On Error Resume Next
  Set getMatches = RE.matchCollection(searchString, regExp)(0).SubMatches
  If Err.Number <> 0 Then
    Err.Clear
    On Error Resume Next
    Set getMatches = RE.matchCollection(searchString, regExp2)(0).SubMatches
    On Error GoTo 0
  End If
End Function
Private Function longCycle(SM As SubMatches) As Boolean
  On Error Resume Next
  longCycle = (SM.Count = 4)
End Function
