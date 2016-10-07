VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cActiveSheetWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Raise one event for any Workbook level close event
'
'Revision History
'
Option Explicit
Private WithEvents Wb  As Workbook
Attribute Wb.VB_VarHelpID = -1
Private WithEvents Ws As Worksheet
Attribute Ws.VB_VarHelpID = -1

Public Event wbBeforeClose(Cancel As Boolean)
Public Event wsLostFocus(Cancel As Boolean)
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim mClosing As Boolean
Private Sub wb_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Const myName As String = "cActiveSheetWrapper.wb_BeforeSave"
Dim db As New cDebugReporter
    db.Report caller:=myName

    If Not mClosing Then
        If SaveAsUI Then
            On Error Resume Next
            Excel.Application.EnableEvents = False
            RaiseEvent wbBeforeClose(Cancel)
            Excel.Application.EnableEvents = True
        End If
    End If
    
End Sub
Private Sub wb_BeforeClose(Cancel As Boolean)
Const myName As String = "cActiveSheetWrapper.wb_BeforeClose"
Dim db As New cDebugReporter
    db.Report caller:=myName

    On Error Resume Next
    Excel.Application.EnableEvents = False
    RaiseEvent wbBeforeClose(Cancel)
    Excel.Application.EnableEvents = True
    
    mClosing = True
    
End Sub
Private Sub Wb_WindowDeactivate(ByVal Wn As Window)
Const myName As String = "cActiveSheetWrapper.Wb_WindowDeactivate"
Dim Cancel As Boolean
Dim db As New cDebugReporter
    db.Report caller:=myName, Context:=Ws.Name

    On Error Resume Next
    Excel.Application.EnableEvents = False
    RaiseEvent wsLostFocus(False)
    Excel.Application.EnableEvents = True

End Sub

Private Sub Ws_Deactivate()
Const myName As String = "cActiveSheetWrapper.Ws_Deactivate"
Dim db As New cDebugReporter
    db.Report caller:=myName, Context:=Ws.Name

    On Error Resume Next
    Excel.Application.EnableEvents = False
    RaiseEvent wsLostFocus(False)
    Excel.Application.EnableEvents = True

End Sub

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Class_Initialize()
Const myName As String = "cActiveSheetWrapper.Class_Initialize"
Dim db As New cDebugReporter
    db.Report caller:=myName

    Set Wb = ActiveWorkbook
    Set Ws = ActiveSheet
    mClosing = False
End Sub

Private Sub Class_Terminate()
Const myName As String = "cActiveSheetWrapper.Class_Terminate"
Dim db As New cDebugReporter
    db.Report caller:=myName, Message:=Wb.Name & vbTab & Ws.Name

    Set Wb = Nothing
    Set Ws = Nothing
    
End Sub

