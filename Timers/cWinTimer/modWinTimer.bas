Attribute VB_Name = "modWinTimer"
Option Explicit
'In modWinTimer
Public Sub TimerProc(ByVal hwnd As Long, _
                     ByVal uMsg As Long, _
                     ByVal oTimer As cWinTimer, _
                     ByVal dwTime As Long)
   ' Alert appropriate timer object instance.
   If Not oTimer Is Nothing Then
     On Error Resume Next
     oTimer.RaiseTimerEvent
     On Error GoTo 0
   End If
End Sub


