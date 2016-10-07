VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWinTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/Usage
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/Private WithEvents mWinTimer  As cWinTimer
'/
'/Private Sub startTimer()
'/
'/    Set mWinTimer = New cWinTimer
'/    mWinTimer.Startit pulseTime
'/
'/End Sub
'/Private Sub stopTimer()
'/
'/    Set mWinTimer = Nothing
'/
'/End Sub
'/
'/Private Sub mWinTimer_OnTimer()
'/
'/    On Error Resume Next
'/    insert code here
'/    On Error GoTo 0
'/
'/End Sub
'/
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/Call-back Must be in standard module
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/In modCallBacks
'/Public Sub TimerProc(ByVal hWnd As Long, _
'/                     ByVal uMsg As Long, _
'/                     ByVal oTimer As cWinTimer, _
'/                     ByVal dwTime As Long)
'/   ' Alert appropriate timer object instance.
'/   If Not oTimer Is Nothing Then
'/     oTimer.RaiseTimerEvent
'/   End If
'/End Sub
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Declare Function SetTimer Lib "user32" ( _
    ByVal hWnd As Long, ByVal nIDEvent As Long, _
    ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" ( _
    ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private mhWnd As LongPtr
Public Event onTimer()
Private mTimerID As Long
Property Let hWnd(RHS As Long)
  mhWnd = RHS
End Property
Property Get hWnd() As Long
  If mhWnd = 0 Then hWnd = getmhWnd Else hWnd = mhWnd
End Property

Property Get TimerID() As Long
  TimerID = mTimerID
End Property
'Start timer
Public Sub Startit(IntervalMs As Long)
Const myName As String = "cWinTimer.Startit"
Dim db As New cDebugReporter
    db.Report caller:=myName


    mTimerID = SetTimer(hWnd, ObjPtr(Me), IntervalMs, AddressOf modCallBacks.TimerProc)
    
    db.ExitMessage = mTimerID & "  " & Hex(hWnd)
    
End Sub

'Stop timer
Public Sub Stopit()
Const myName As String = "cWinTimer.Stopit"
Dim db As New cDebugReporter
    db.Report caller:=myName

    If TimerID <> -1 Then
        KillTimer hWnd, TimerID
        mTimerID = 0
    End If
End Sub


'Trigger Public event
Public Sub RaiseTimerEvent()
Const myName As String = "cWinTimer.RaiseTimerEvent"

    RaiseEvent onTimer
End Sub


Private Sub Class_Initialize()
Const myName As String = "cWinTimer.Class_Initialize"
Dim db As New cDebugReporter
    db.Report caller:=myName

End Sub

Private Sub Class_Terminate()
Const myName As String = "cWinTimer.Class_Terminate"
Dim db As New cDebugReporter
    db.Report caller:=myName

  Stopit
  
End Sub

Private Function getmhWnd() As Long
Dim mPPTWindow As cpptWindow
Dim hWnd As Long
  On Error GoTo notAvailable
  hWnd = CallByName(Application, "HWnd", VbGet)
  
  getmhWnd = hWnd
  Exit Function
notAvailable:
  Set mPPTWindow = New cpptWindow
  'hWnd = mPPTWindow.HandleByClass(ActiveWindow, "screenClass")
  If hWnd = 0 Then
  ' reset the object to get the default class name
    Set mPPTWindow = New cpptWindow
    hWnd = mPPTWindow.Handle(ActiveWindow)
  End If
  Set mPPTWindow = Nothing
  Resume Next
End Function
