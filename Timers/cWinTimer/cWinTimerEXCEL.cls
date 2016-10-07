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
'/    alignWithWindow
'/    On Error GoTo 0
'/
'/End Sub
'/
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/Call-back
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/In modWinTimer
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
    ByVal hwnd As Long, ByVal nIDEvent As Long, _
    ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private mhWnd As LongPtr
Public Event OnTimer()

Private TimerID As Long
'Start timer
Public Sub Startit(IntervalMs As Long)
    mhWnd = Application.hwnd
    TimerID = SetTimer(mhWnd, ObjPtr(Me), IntervalMs, AddressOf modWinTimer.TimerProc)
End Sub

'Stop timer
Public Sub Stopit()
    If TimerID <> -1 Then
        KillTimer mhWnd, TimerID
        TimerID = 0
    End If
End Sub


'Trigger Public event
Public Sub RaiseTimerEvent()
    RaiseEvent OnTimer
End Sub


Private Sub Class_Terminate()
    Stopit
End Sub
