VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAnnimateFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////////////
' Module    : cAnnimateFrames
' Author    : cool blue cool.blue@y7mail.com
' Date      : 29/09/2014
' Purpose   : Animate an arbitrary property of the object wrapped in mShapeByName
'//////////////////////////////////////////////////////////////////////////////////////
Const CmoduleName As String = "Class Module cAnnimateFrames"
Option Explicit

Private WithEvents mWinTimer  As cWinTimer
Attribute mWinTimer.VB_VarHelpID = -1

Public Event annComplete()

Private mShapeByName As cObjectByName

Public startValue As Integer
Public EndValue As Integer
Public StepDirection As Integer
Public dt As Long
Public T As Integer
Private v As Integer
Private stepCount As Integer
Private mEaseOut As cannEaseOut
Private testContinue As Boolean

Private mTriggerText As Shape
Property Let TriggerText(RHS As Shape)
  Set mTriggerText = RHS
End Property
Property Let ShapeByName(RHS As cObjectByName)
  Set mShapeByName = RHS
End Property
Private Sub startTimer(dt As Long)
Dim mPPTWindow As cpptWindow
  Set mWinTimer = New cWinTimer
  Set mPPTWindow = New cpptWindow
  With mWinTimer
    .hWnd = mPPTWindow.Handle(ActiveWindow)
    .Startit dt
  End With
End Sub
Private Sub stopTimer()
  Set mWinTimer = Nothing
  Set mEaseOut = Nothing
  testContinue = False
End Sub
Public Sub Start()
  timerStart
End Sub

Private Sub timerStart()
  Set mEaseOut = New cannEaseOut
  With mEaseOut
    .dt = dt: .startValue = startValue: .EndValue = EndValue: .T = T
  End With
  v = startValue
  stepCount = 0
'  DoEvents
  startTimer dt
  mWinTimer_onTimer
End Sub
Private Sub testStart()
  Set mEaseOut = New cannEaseOut
  With mEaseOut
    .dt = dt: .startValue = startValue: .EndValue = EndValue: .T = T
  End With
  v = startValue
  stepCount = 0
  testContinue = True
  Do
    mWinTimer_onTimer
    DoEvents
  Loop While testContinue
End Sub

Private Sub mWinTimer_onTimer()
Const myName As String = "cAnnimateFrames.mWinTimer_onTimer"
Dim StepValue As Integer
Dim db As New cDebugReporter
    db.Report caller:=myName ', Message:=mWinTimer.TimerID
  
  With mShapeByName
    .Value = CInt(v)
    db.Report Message:="read-back value: " & .Value
    StepValue = mEaseOut.dCosT()
    v = v + StepValue
    If (EndValue - v) * StepDirection <= 0 Then
      db.Report Context:="end loop", Message:=v & vbTab & StepValue
      stopTimer
      .Value = EndValue
      RaiseEvent annComplete
    End If
  End With
  With mTriggerText
    .TextFrame.TextRange.Text = .TextFrame.TextRange.Text
  End With
  stepCount = stepCount + 1
  db.ExitMessage = v
End Sub
Private Sub Class_Initialize()
Const myName As String = "cAnnimateFrames.Class_Initialize"
Dim db As New cDebugReporter
    db.Report caller:=myName

End Sub
Private Sub Class_Terminate()
Const myName As String = "cAnnimateFrames.Class_Terminate"
Dim db As New cDebugReporter
    db.Report caller:=myName

End Sub
