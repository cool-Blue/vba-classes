VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cActiveSheetBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Notify Activesheet of Workbook Events
'Revision History
'   [20131127]
'   Added beforeSaveAs Call-back

Option Explicit
Const cModuleName = "cActiveSheetBus"
Const cmIndent = 1
Sub activeSheetCallBack(Ws As Worksheet, cb As String)
Const myName As String = "cActiveSheetBus.activeSheetCallBack"
    
    On Error GoTo fnCallbackFailed
    CallByName Ws, cb, VbMethod
    On Error GoTo 0
    Exit Sub
fnCallbackFailed:
    Debug.Print cModuleName & vbTab & myName & vbTab & "****failed****"
    Err.Clear
End Sub

Public Sub onOpen(Ws As Worksheet)
    activeSheetCallBack Ws, "onOpen"
End Sub
Public Sub beforeClose(Ws As Worksheet)
    activeSheetCallBack Ws, "beforeClose"
End Sub
'[20131127]
Public Sub beforeSaveAs(Ws As Worksheet)
    activeSheetCallBack Ws, "beforeSaveAs"
End Sub
Public Sub beforeSave(Ws As Worksheet)
    activeSheetCallBack Ws, "beforeSave"
End Sub
Public Sub afterSave(Ws As Worksheet)
    activeSheetCallBack Ws, "afterSave"
End Sub
Public Sub onWindowActivate(Ws As Worksheet)
    activeSheetCallBack Ws, "onWindowActivate"
End Sub
Public Sub onWindowDEActivate(Ws As Worksheet)
    activeSheetCallBack Ws, "onWindowDEActivate"
End Sub

Private Sub Class_Terminate()
Const myName As String = "cActiveSheetBus.Class_Terminate"
    Debug.Print cModuleName & vbTab & myName
End Sub
Private Sub Class_Initialize()
Const myName As String = "cActiveSheetBus.Class_Initialize"
    Debug.Print cModuleName & vbTab & myName
End Sub
