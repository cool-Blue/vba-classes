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
Sub activeSheetCallBack(ws As Worksheet, cb As String)
Const myName As String = "activeSheetCallBack"
    
    On Error GoTo fnCallbackFailed
    If debugMode Then Debug.Print timeStamp(Indent:=cmIndent, Caller:=cModuleName, Context:=myName, message:="call " & cb & " in sheet: " & ws.Name)
    CallByName ws, cb, VbMethod
    On Error GoTo 0
    Exit Sub
fnCallbackFailed:
    If debugMode Then Debug.Print timeStamp(Indent:=cmIndent, Caller:=cModuleName, Context:=myName, message:=cb & " FAILED in sheet: " & ws.Name)
    Err.Clear
End Sub

Public Sub onOpen(ws As Worksheet)
    activeSheetCallBack ws, "onOpen"
End Sub
Public Sub beforeClose(ws As Worksheet)
    activeSheetCallBack ws, "beforeClose"
End Sub
'[20131127]
Public Sub beforeSaveAs(ws As Worksheet)
    activeSheetCallBack ws, "beforeSaveAs"
End Sub
Public Sub beforeSave(ws As Worksheet)
    activeSheetCallBack ws, "beforeSave"
End Sub
Public Sub afterSave(ws As Worksheet)
    activeSheetCallBack ws, "afterSave"
End Sub
Public Sub onWindowActivate(ws As Worksheet)
    activeSheetCallBack ws, "onWindowActivate"
End Sub
Public Sub onWindowDEActivate(ws As Worksheet)
    activeSheetCallBack ws, "onWindowDEActivate"
End Sub

Private Sub Class_Terminate()
Const cpName As String = "Class_Terminate"

    If debugEvents Then Debug.Print timeStamp(Indent:=cmIndent, Caller:=cModuleName, Context:=cpName)
    
End Sub
Private Sub Class_Initialize()
Const cpName As String = "Class_Initialize"

    If debugEvents Then Debug.Print timeStamp(Indent:=cmIndent, Caller:=cModuleName, Context:=cpName)
End Sub
