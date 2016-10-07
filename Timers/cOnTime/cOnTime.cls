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
'Encapsulated timer that will sense the active
' sheet and expect to find a callback there
'
'In host sheet
' Const cPulseTime As Long = 1
'
' Dim mOnTime As cOnTime
' Property Get PulseTime() As Long
'     PulseTime = cPulseTime
' End Property
' '****************************************
' 'Timer call-back for cOnTime
' Public Sub kickDog()
' '   Code to execute on timer event
' '******************************************
'     On Error Resume Next
'     Me.Cells(1, 1) = Not Me.Cells(1, 1)
'     On Error GoTo 0
' '******************************************
'     Debug.Print "woof!!"
'     On Error GoTo exitError
'         mOnTime.kickDog
'     On Error GoTo 0
' Exit Sub
' exitError:
' End Sub

Const DEFDoWhen As String = "kickDog"
Const DEFPulseTime = "PulseTime"
Const DEFearliestTime As Long = 1
Const DEFlatestTime As Long = 15

Private WithEvents wb As Workbook
Attribute wb.VB_VarHelpID = -1
Private Ws As Worksheet
Attribute Ws.VB_VarHelpID = -1

Private DoWhen As String
Dim KillTimer As String
Private mPulseTime As Long
Private mDesignMode
Private mBackgroundPulse
Private mNextTime As Double
Property Let callBackDoWhen(cb As String)
    DoWhen = "'" & wb.Name & "'!" & Ws.CodeName & "." & cb      'e.g. 'wb Name.xlsm'!Sheet1.kickdog
End Property
Property Let callBackPulseTime(csPulseTime As String)
Const cMyName As String = "Let PulseTime"

    On Error Resume Next
         mPulseTime = CallByName(Ws, csPulseTime, VbGet)
         If Err.Number <> 0 Then
             mPulseTime = DEFearliestTime
         End If
    On Error GoTo 0

End Property
Private Function wsGetProperty(prop As String, default)
    On Error Resume Next
         wsGetProperty = CallByName(Ws, prop, VbGet)
         If Err.Number <> 0 Then
             wsGetProperty = default
         End If
    On Error GoTo 0
End Function
Private Function pulseTime() As Long
' This is a live connection to the sheet
    pulseTime = wsGetProperty(DEFPulseTime, DEFearliestTime)
End Function
Private Function designMode() As Boolean
' The sheet is only consulted once
  If mDesignMode = Empty Then _
    mDesignMode = wsGetProperty("designMode", False)
  designMode = mDesignMode
End Function
Private Function backgroundPulse() As Boolean
' The sheet is only consulted once
  If mBackgroundPulse = Empty Then _
    mBackgroundPulse = wsGetProperty("enableBackgroundPulse", False)
  backgroundPulse = mBackgroundPulse
End Function
Public Sub kickDog()
Const myName As String = "kickDog"
Dim psMessage As String
    If (Ws Is ActiveSheet Or backgroundPulse) _
        And Not designMode Then
        
        mNextTime = Now + TimeSerial(0, 0, pulseTime)
        
        On Error Resume Next
        Application.OnTime mNextTime, DoWhen
        On Error GoTo 0
    End If
    Exit Sub
End Sub
Public Sub killDog()

    If Ws Is Nothing Or mNextTime = 0 Then Exit Sub
        
    On Error Resume Next
    Application.OnTime mNextTime, DoWhen, , False
    On Error GoTo 0

End Sub
Private Sub Class_Initialize()
Dim errorContext As String

Debug.Print "init conTime"

    On Error GoTo enableAndExit
        Set wb = ActiveWorkbook
        Set Ws = ActiveSheet
    On Error GoTo 0
    
    callBackDoWhen = DEFDoWhen
    callBackPulseTime = DEFPulseTime
    pulseTime
    designMode
    backgroundPulse
    
    kickDog
    
    Exit Sub
enableAndExit:
    If Err <> 0 Then
        If Ws Is Nothing Then
            errorContext = "ws"
        ElseIf wb Is Nothing Then
            errorContext = "wb"
        End If
    End If
End Sub

Private Sub Class_Terminate()
Const myName As String = "Class_Terminate"

    On Error Resume Next
    killDog
    Set Ws = Nothing
    Set wb = Nothing
    Exit Sub

End Sub

' Manage the timer in response to workbook events
' If the timer is not killed it may cause the workbook
' to reopen after it is closed when the timer calls back.

Private Sub wb_WindowActivate(ByVal Wn As Window)
Const myName As String = "cOnTime.wb_WindowActivate"
    Debug.Print myName

   ' this is handled by ThisWorkbook
End Sub

Private Sub wb_WindowDeactivate(ByVal Wn As Window)
Const myName As String = "cOnTime.wb_WindowDeactivate"
    Debug.Print myName

    If Not backgroundPulse Then killDog
End Sub
Private Sub wb_BeforeClose(Cancel As Boolean)
Const myName As String = "cOnTime.wb_BeforeClose"
    Debug.Print myName

    killDog
End Sub

Private Sub wb_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Const myName As String = "cOnTime.wb_BeforeSave"
    Debug.Print myName

    If SaveAsUI Then killDog
End Sub


