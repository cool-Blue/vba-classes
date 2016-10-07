VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'In Worksheet Module

Const cPulseTime As Long = 1
Const cBackgroundPulse As Boolean = False
Dim mOnTime As cOnTime

'Expose custom worksheet propertiesto configurethe timer
Property Get pulseTime() As Long
  ' Can put any logic here that interracts with the sheet
  ' or the user or the application for example
  '  pulseTime = cPulseTime
    pulseTime = Me.Range("pulseTime")
End Property
Property Get enableBackgroundPulse() As Boolean
    enableBackgroundPulse = cBackgroundPulse
End Property
Property Get designMode() As Boolean
    designMode = Me.ProtectContents
End Property

'****************************************
'ActiveSheet Call-backs
Public Sub onWindowActivate()
Const cMyName As String = "onWindowActivate"

  Worksheet_Activate

End Sub
'****************************************

'****************************************
'Timer call-back for cOnTime
Public Sub kickDog()
'   Code to execute on timer event
'******************************************
    On Error Resume Next
    Me.Cells(1, 1) = Not Me.Cells(1, 1)
    On Error GoTo 0
'******************************************
    Debug.Print "woof!!"
    On Error GoTo exitError
        mOnTime.kickDog
    On Error GoTo 0
Exit Sub
exitError:
End Sub

Private Sub Worksheet_Activate()
Const myName As String = "Sheet2.Worksheet_Activate"
    Debug.Print myName
    
    If (mOnTime Is Nothing) Then
      Set mOnTime = New cOnTime
    Else
      mOnTime.kickDog
    End If
End Sub

Private Sub Worksheet_Deactivate()
Const pName As String = "Sheet2.Worksheet_Deactivate"

  

End Sub
