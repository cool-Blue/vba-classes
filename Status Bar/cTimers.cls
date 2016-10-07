VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mTimers As Collection

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/ common call back
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fireTimer(timer As String)
Const myName As String = "cTimers.fireTimer"
Dim db As New cDebugReporter
    db.Report Caller:=myName

  ' Alert appropriate timer object instance.
  On Error Resume Next
  mTimers(timer).RaiseTimerEvent
  
End Sub
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/ collection management
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function Add() As cOnTimer
  Set Add = New cOnTimer
  Add.Name = ObjPtr(Add)
  If mTimers Is Nothing Then Set mTimers = New Collection
  On Error Resume Next
  mTimers.Add Add, Add.Name
End Function
Public Function Delete(timer As cOnTimer)
    timer.Stopit
    With mTimers
      .Remove timer.Name
      If .Count = 0 Then Set mTimers = Nothing
    End With
    Set timer = Nothing
End Function

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/ class life cycle management
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
  Set mTimers = New Collection
End Sub

Public Sub Terminate()
Const myName As String = "cTimers.Terminate"
Dim db As New cDebugReporter
    db.Report Caller:=myName

Dim oTimer As Object
  If Not mTimers Is Nothing Then
    For Each oTimer In mTimers
      oTimer.Stopit
      mTimers.Remove oTimer.Name
    Next oTimer
  End If
End Sub

