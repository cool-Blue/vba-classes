VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cImediateWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/ Module for clearing the imediate window
'/ http://stackoverflow.com/questions/10203349/use-vba-to-clear-immediate-window
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Declare Function GetWindow Lib "user32" ( _
                          ByVal hWnd As Long, _
                          ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                          ByVal lpClassName As String, _
                          ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
                          (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
                          ByVal lpsz1 As String, _
                          ByVal lpsz2 As String) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
                          ByVal hWnd As Long, ByVal wMsg As Long, _
                          ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                          ByVal hWnd As Long, ByVal wMsg As Long, _
                          ByVal wParam As Long, ByVal lParam As Long) As Long


Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function getWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
 
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_SETFOCUS As Long = &H7

Private Const KEYSTATE_KEYDOWN As Long = &H80
'Private savState(0 To 255) As Byte

Public Sub ClearTrace()
Const myName As String = "cDebugReporter.ClearTrace"
Dim db As New cDebugReporter
    db.Report caller:=myName

  ClearImmediateWindow
'  DoEvents
'  DoCleanUp
  db.Report Message:="TRACE CLEARED"
End Sub
Private Sub ClearImmediateWindow()
Const myName As String = "cDebugReporter.ClearImmediateWindow"
Dim db As New cDebugReporter
    db.Report caller:=myName

'Adapted  by   keepITcool
'Original from Jamie Collins fka "OneDayWhen"
'http://www.dicks-blog.com/excel/2004/06/clear_the_immed.html

Dim hPane As Long, imWnd As Object, hcpPane As Long, cpWnd As Object
Dim tmpState(0 To 255) As Byte

  With Application.VBE.MainWindow
    .SetFocus
    .Visible = True
  End With

  hcpPane = GetWinHandle(cpWnd, vbext_wt_CodeWindow)
  hPane = GetWinHandle(imWnd, vbext_wt_Immediate)
  db.Report Message:="hPane: 0x" & Hex(hPane)
  
  If hPane = 0 Then MsgBox "Immediate Window not found."
  If hPane < 1 Then Exit Sub
  'Save the keyboardstate

  GetKeyboardState gSavState(0)
  PostMessage hPane, WM_SETFOCUS, 0&, 0&
'    SendMessage hPane, WM_SETFOCUS, 0&, 0&
  
  'Send CTRL-A
    'Sink the CTRL (note we work with the empty tmpState)
    tmpState(vbKeyControl) = KEYSTATE_KEYDOWN
    SetKeyboardState tmpState(0)
'    PostMessage hPane, WM_KEYDOWN, vbKeyControl, &H1
    PostMessage hPane, WM_KEYDOWN, vbKeyA, &H1
    PostMessage hPane, WM_KEYUP, vbKeyA, &HC0000001
    
  'Send DELETE
    PostMessage hPane, WM_KEYDOWN, vbKeyDelete, &H1
    PostMessage hPane, WM_KEYUP, vbKeyDelete, &HC0000001
    
  db.Report Message:=imWnd.Caption
'  imWnd.Close
'  Application.OnTime Now + TimeSerial(0, 0, 0), "gDoCleanUp"
End Sub
Private Function GetWinHandle(Optional ByRef oWnd As Object, _
                              Optional windowType As Long, _
                              Optional parentWnd As Long) As Long
Const myName As String = "cDebugReporter.GetImmHandle"
Const VBE_WINDOW_CLASS_NAME As String = "wndclass_desked_gsk"
Dim db As New cDebugReporter
    db.Report caller:=myName

'This function finds the Immediate Pane and returns a handle.
'Docked or MDI, Desked or Floating, Visible or Hidden

Dim bDock As Boolean, bShow As Boolean
Dim sMain$, sDock$, sPane$
Dim lMain&, lDock&, lPane&
Dim isFound As Boolean, str As String

  On Error Resume Next
' Get the main VBE window Caption
  sMain = Application.VBE.MainWindow.Caption
  If Err <> 0 Then
    MsgBox "No Access to Visual Basic Project"
    GetWinHandle = -1
    Exit Function
    ' Excel2003: Registry Editor (Regedit.exe)
    '    HKLM\SOFTWARE\Microsoft\Office\11.0\Excel\Security
    '    Change or add a DWORD called 'AccessVBOM', set to 1
    ' Excel2002: Tools/Macro/Security
    '    Tab 'Trusted Sources', Check 'Trust access..'
  End If
  
' Find the imediate window and log its properties
  For Each oWnd In Application.VBE.Windows
    If oWnd.Type = windowType Then
      bShow = oWnd.Visible
      sPane = oWnd.Caption
      If Not oWnd.LinkedWindowFrame Is Nothing Then
        bDock = True
        sDock = oWnd.LinkedWindowFrame.Caption
      End If
      Exit For
    End If
  Next
' Get the main Window hWnd
  lMain = FindWindow(VBE_WINDOW_CLASS_NAME, sMain)
  If lMain = 0 Then
    db.Report Message:=sMain & " Length: " & Len(sMain)
  ' Caption could be truncated or have added ornamentation...
  ' Find a window who's caption contains the target caption.
    lMain = FindWindowEx(0&, 0&, VBE_WINDOW_CLASS_NAME, vbNullString)
    Do While lMain <> 0 And Not isFound
        str = getWindowTextHWnd(lMain)
        If InStr(str, sMain) Then
            isFound = True
            Exit Do
        Else
            lMain = FindWindowEx(0&, lMain, VBE_WINDOW_CLASS_NAME, vbNullString)
        End If
    Loop
  End If
' Search on the Main window to find the immediate window
' using its caption.
  If bDock Then
  'Docked within the VBE
    lPane = FindWindowEx(lMain, 0&, "VbaWindow", sPane)
    If lPane = 0 Then
      'Floating Pane.. which MAY have it's own frame
      'Assume FindWindowEx does not loop back to the first window if no matching windows
      'are found and that it returns zero after it tries to match the last of the class.
      'Also assume that there is only one VbaWindow object for each VbFloatingPalette.
      Do
        lDock = FindWindowEx(0&, lDock, "VbFloatingPalette", vbNullString)
        lPane = FindWindowEx(lDock, 0&, "VbaWindow", sPane)
      Loop While lDock > 0 And lPane = 0
      
    End If
  ElseIf bShow Then
    lDock = FindWindowEx(lMain, 0&, "MDIClient", _
    vbNullString)
    lDock = FindWindowEx(lDock, 0&, "DockingView", _
    vbNullString)
    If lDock = 0 Then
      lDock = FindWindowEx(lMain, 0&, "MDIClient", vbNullString)
      lPane = FindWindowEx(lDock, 0&, "VbaWindow", sPane)
    Else
      lPane = FindWindowEx(lDock, 0&, "VbaWindow", sPane)
    End If
  Else
    lPane = FindWindowEx(lMain, 0&, "VbaWindow", sPane)
  End If
  
  GetWinHandle = lPane
  parentWnd = lDock

End Function
Private Function getWindowTextHWnd(hWnd As LongPtr) As String
Dim txt As String, txtLen As Long

    txtLen = GetWindowTextLength(hWnd) + 1
    txt = String(txtLen, vbNullChar)
    Call getWindowText(hWnd, txt, txtLen)
    getWindowTextHWnd = txt
    
End Function


