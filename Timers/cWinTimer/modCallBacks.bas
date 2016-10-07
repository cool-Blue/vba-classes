Attribute VB_Name = "modCallBacks"
Option Explicit

Public Sub TimerProc(ByVal hWnd As Long, _
                     ByVal uMsg As Long, _
                     ByVal oTimer As cWinTimer, _
                     ByVal dwTime As Long)
Const myName As String = "modCallBacks.TimerProc"

   ' Alert appropriate timer object instance.
   If Not oTimer Is Nothing Then
     On Error Resume Next
     
'    onTimer
     oTimer.RaiseTimerEvent
'     oTimer.Stopit
     On Error GoTo 0
   End If
End Sub
Public Sub EnumChildProc(ByVal hWnd As Long, _
                         ByVal ohWnd As cHWnd)
  
End Sub


