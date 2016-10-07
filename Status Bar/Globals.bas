Attribute VB_Name = "Globals"
Option Explicit
Public gnextTime As Date

Public Const gcDebugMode As Boolean = True
Public Const gdebugOutoutToFile As Boolean = True
Public gCallDepth As Long
Public glogFile As cTextStream
Public gSavState(0 To 255) As Byte

Private Sub testEvents()
Const myName As String = "Globals.testEvents"
Dim db As New cDebugReporter
    db.Report Caller:=myName

Dim parentSubscriber As cParentSubscriber

  Set parentSubscriber = New cParentSubscriber
  db.Report message:="setting to nothing"
  Set parentSubscriber = Nothing
  db.Report message:="set to nothing"
  
End Sub
Public Sub testCallBack(name As String, nextTime As String)
  MsgBox "callback " & name & " " & nextTime
End Sub
Private Sub testTimerSet()
Const myName As String = "Globals.testTimerSet"
Dim db As New cDebugReporter
    db.Report Caller:=myName

  gnextTime = Now() + TimeSerial(1, 0, 0)
  Application.OnTime gnextTime, sProcedure("Globals.testCallBack", _
                                            "testTimer", gnextTime)
                                            
  db.Report message:=sProcedure("Globals.testCallBack", _
                                            "testTimer", gnextTime)
End Sub
Public Sub testTimerKill()
Const myName As String = "Globals.testTimerKill"
Dim db As New cDebugReporter
    db.Report Caller:=myName

  On Error Resume Next
  Application.OnTime gnextTime, sProcedure("Globals.testCallBack", _
                                            "testTimer", gnextTime), , False
End Sub

Public Sub testClose()
Const myName As String = "Globals.testClose"
Dim db As New cDebugReporter
    db.Report Caller:=myName

  Application.OnTime Now(), "Globals.closeWorkbook"
  
  db.Report message:="shutDown Scheduled"
End Sub
Public Sub closeWorkbook()
Const myName As String = "ThisWorkbook.closeWorkbook"
Dim db As New cDebugReporter
    db.Report Caller:=myName

    With ThisWorkbook
        .Saved = True
        db.Report message:="Saved"
        .Close
        db.Report message:="Closed"
    End With

End Sub


Private Function fmtTime(d As Date) As String
  fmtTime = Format(Hour(d), "00") & ":" & Format(Minute(d), "00") & ":" & Format(Second(d), "00")
End Function
Private Function sProcedure(callBackProcedure As String, mName As String, nextTime As Date) As String
  sProcedure = "'" & callBackProcedure & " " & """" & mName & """," & """" & fmtTime(nextTime) & """'"
End Function

