VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOnTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'****************************************
'Host Sheet call-back (MUST BE CALLED FROM Active Sheet):
'Timer call-back for cOnTime
'Public Sub kickDog()
'    ws.kickDog
'End Sub
'Private Sub killTimer()
'    Set mOnTime = Nothing
'End Sub
'****************************************
'Revision History
'20131125
'   added Property Get earliestTime() As Long
'   globalised mPulseTime
'   removed newline from terminate report

'Const debugMode = True

Const g1earliestTime As Long = 5

Public Event killTimerFail(doWhen As String)
Private Wb As Workbook
Private Ws As Worksheet

Private mPulseTime As Long

Public doWhen As String
Public NextTime As Double
Public startTime As Double
Property Get earliestTime() As Long
    earliestTime = g1earliestTime
End Property
Property Get pulseTime() As Long
    pulseTime = mPulseTime
End Property
Property Let pulseTime(LHS As Long)
    mPulseTime = LHS
End Property
Property Get designMode() As Boolean
    designMode = Not Ws.ProtectContents
End Property
Public Sub kickDog()
Const myName As String = "cOnTime.kickDog"
Dim db As New cDebugReporter
Dim psMessage As String
    db.Report caller:=myName

    If Ws Is ActiveSheet And Not designMode Then
        
'       Set the timing for the next pulse
        NextTime = Now + TimeSerial(0, 0, mPulseTime)
        
'       Report
        db.Report Message:=doWhen
        db.Report Message:="NextTime: " & Text(NextTime, 10)
        db.Report Message:="PulseTime: " & mPulseTime
        db.Report Message:="t + " & IIf(startTime <> 0, Text(MicroTimer - startTime, 2), 0)
        startTime = MicroTimer
        
'       Restart Timer
        On Error Resume Next
        Excel.Application.OnTime NextTime, doWhen
        If Err <> 0 Then
            db.ExitMessage = doWhen
            Err.Clear
        End If
        
        On Error GoTo 0
    End If
    Exit Sub
End Sub
Public Sub killDog()
Const myName As String = "cOnTime.killDog"
Dim doWhenRefreshed As Boolean
Dim db As New cDebugReporter
    db.Report caller:=myName

    If Ws Is Nothing Or NextTime = 0 Then Exit Sub
    
    db.Report Message:=doWhen
    db.Report Message:="NextTime: " & Text(NextTime, 10)
    db.Report Message:="t + " & Text(MicroTimer - startTime, 2)
    
    On Error Resume Next
    Excel.Application.OnTime NextTime, doWhen, , False
    If Err <> 0 Then
        If Not doWhenRefreshed Then
            RaiseEvent killTimerFail(doWhen)
            doWhenRefreshed = True
            db.Report Message:="doWhen Refreshed"
            Resume
        Else
            db.ExitMessage = "FAILED: " & doWhen
            Err.Clear
        End If
    End If
    
    On Error GoTo 0
    Exit Sub

End Sub
Private Sub Class_Initialize()
Const myName As String = "cOnTime.Class_Initialize"
Dim db As New cDebugReporter
Dim errorContext As String

    db.Report caller:=myName

    On Error GoTo enableAndExit
        Set Wb = ActiveWorkbook
        Set Ws = ActiveSheet
        db.Report Message:=Ws.Name
    On Error GoTo 0
    Exit Sub
enableAndExit:
    If Err <> 0 Then
        If debugErrorMode Then
            If Ws Is Nothing Then
                errorContext = "ws"
            ElseIf Wb Is Nothing Then
                errorContext = "wb"
            End If
            db.ExitMessage = errorContext
        End If
    End If
End Sub

Private Sub Class_Terminate()
Const myName As String = "cOnTime.Class_Terminate"
Dim db As New cDebugReporter
    db.Report caller:=myName


    On Error Resume Next
    db.Report Message:=doWhen
    killDog
    Set Ws = Nothing
    Set Wb = Nothing
    Exit Sub
    If Err <> 0 Then
            db.ExitMessage = doWhen
    End If
End Sub



