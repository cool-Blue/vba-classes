Attribute VB_Name = "modTestTimers"
Option Explicit

Public gnextTime As Date

Public Sub testCallBack(name As String, nextTime As String)
  MsgBox "callback " & name & " " & nextTime
End Sub
Private Sub testTimerSet()
  gnextTime = Now() + TimeSerial(1, 0, 0)
  Application.OnTime gnextTime, sProcedure("modTestTimers.testCallBack", _
                                            "testTimer", gnextTime)
    Debug.Print sProcedure("modTestTimers.testCallBack", "testTimer", gnextTime), "Set"

End Sub
Public Sub testTimerKill()
Const myName As String = "testTimerKill"
  Debug.Print myName, "START"
  
  On Error Resume Next
  Application.OnTime gnextTime, sProcedure("modTestTimers.testCallBack", _
                                            "testTimer", gnextTime), _
                                            , False
  If Err <> 0 Then
    Debug.Print Err, Err.Description
  Else
    Debug.Print "No Error killing " & sProcedure("modTestTimers.testCallBack", _
                                            "testTimer", gnextTime)
  End If
  
  Debug.Print myName, "END"
End Sub

Public Sub testClose()
Const myName As String = "testClose"
  Debug.Print myName, "START"
  
  Application.OnTime Now(), "modTestTimers.closeWorkbook"
  Debug.Print "shutDown Scheduled"
  
  Debug.Print myName, "END"
End Sub
Public Sub closeWorkbook()
Const myName As String = "closeWorkbook"
  Debug.Print myName, "START"

    With ThisWorkbook
        .Saved = True
        Debug.Print myName, "Closing..."
        .Close
    End With
  Debug.Print myName, "END"
End Sub

Private Function fmtTime(d As Date) As String
  fmtTime = Format(Hour(d), "00") & ":" & Format(Minute(d), "00") & ":" & Format(Second(d), "00")
End Function
Private Function sProcedure(callBackProcedure As String, mName As String, nextTime As Date) As String
  sProcedure = "'" & callBackProcedure & " " & """" & mName & """," & """" & fmtTime(nextTime) & """'"
End Function

