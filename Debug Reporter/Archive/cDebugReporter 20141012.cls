VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDebugReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/ Maintains a global call stack
'/ Requires
'/ Public const gcDebugMode as boolean = ???
'/ Public const gdebugOutoutToFile as Boolean = ???
'/ Public gCallDepth as Long
'/ Public glogFile As cTextStream
'/
'/ Revision History
'/ 29/09/2014
'/   Generalised to work without Excel Application calls
'/   Now works in Powerpoint
'/   Eliminated all outboard dependencies: now independent
'/   Error messages put on a seperate line
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Enum eNewLine
    No
    Before
    After
    Both
    Auto
End Enum

Private mLevel As Long
Private mFirstReport As Boolean
Private startTime As Double, lapTime As Double
Private mCaller As String
Private mContext As String
Public ExitMessage As String

Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
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
  Report Message:="TRACE CLEARED"
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
Private Sub ClearImmediateWindowOLD()
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

  With imWnd
    .SetFocus
    .Visible = True
  End With
  
  GetKeyboardState gSavState(0)
  'Send CTRL+End
    'Sink the CTRL (note we work with the empty tmpState)
    tmpState(vbKeyControl) = KEYSTATE_KEYDOWN
    SetKeyboardState tmpState(0)
    'Sink End
    PostMessage hPane, WM_KEYDOWN, vbKeyEnd, 0&
    PostMessage hPane, WM_KEYUP, vbKeyEnd, 0&
  'db.Report Message:="*******************Wake up!!*******************"
  'Send CTRLSHIFT+Home and CTRLSHIFT+BackSpace
    'Sink the SHIFT
    tmpState(vbKeyShift) = KEYSTATE_KEYDOWN
    SetKeyboardState tmpState(0)
    'Sink Home key
    PostMessage hPane, WM_KEYDOWN, vbKeyHome, 0&
    PostMessage hPane, WM_KEYUP, vbKeyHome, 0&
    'Release control and shift keys
'    SetKeyboardState savState(0) ***** this breaks it in BREAK Mode
    'Sink Back key
    PostMessage hPane, WM_KEYDOWN, vbKeyBack, 0&
    PostMessage hPane, WM_KEYUP, vbKeyBack, 0&
  db.Report Message:=imWnd.Caption
'  imWnd.Close
  
End Sub

Public Sub DoCleanUp(savState As Byte)
Const myName As String = "cDebugReporter.DoCleanUp"
Dim db As New cDebugReporter
    db.Report caller:=myName
' Restore keyboard state
  SetKeyboardState savState
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

Public Function max(va, vb) As Variant
  If va > vb Then max = va Else max = vb
End Function
Public Function sysTime() As String
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim MyTime As SYSTEMTIME
    'Set the graphical mode to persistent
    'Get the local time
    GetLocalTime MyTime
    'Print it to the form
'    On Error GoTo exitEarly
    On Error Resume Next
    sysTime = Format(MyTime.wHour, "00") & ":" & Format(MyTime.wMinute, "00") & ":" & _
                Format(MyTime.wSecond, "00") & ":" & Format(MyTime.wMilliseconds, "000")
exitEarly:
End Function

Public Function timeStamp(Optional d As Double = 0, Optional newLine As eNewLine = No, Optional Indent As Long = 0, _
                            Optional caller As String, Optional Context As String, Optional Message As String, _
                            Optional dt As String, Optional dtCol As Long = 100) As String
Dim errorMessage As String
    If Err.Number <> 0 Then
        errorMessage = "ERROR: " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    
    If d = 0 Then
        timeStamp = sysTime & String(Indent, " ")
    Else
      timeStamp = Format(Hour(d), "00") & ":" & Format(Minute(d), "00") & ":" & Format(Second(d), "00") & ":" & String(Indent, " ")
    End If
    If Len(caller) <> 0 Then timeStamp = timeStamp & caller
    If Len(Context) <> 0 Then timeStamp = timeStamp & ": " & " " & Context
    If Len(Message) <> 0 Then timeStamp = timeStamp & IIf(Len(Context) = 0, ": ", "  ") & " " & Message
    If Len(errorMessage) <> 0 Then
        timeStamp = timeStamp & vbNewLine & sysTime & String(Indent, "*") & errorMessage
    End If
    If Len(dt) <> 0 Then
        timeStamp = timeStamp & setCol(timeStamp, dt, dtCol)
    End If
    Select Case newLine
    Case Before
        timeStamp = vbNewLine & timeStamp
    Case After
        timeStamp = timeStamp & vbNewLine
    Case Both
        timeStamp = vbNewLine & timeStamp & vbNewLine
    Case Else
    End Select
    
End Function
Function setCol(S1 As String, S2 As String, s2Col As Long, Optional divider As String = ":") As String
    
    If s2Col < 0 Then
        setCol = " " & divider & S2
    Else
        setCol = String(max(s2Col - Len(S1) - Len(divider), 0), " ") & divider & S2
    End If
    
End Function

Public Sub Report(Optional d As Double = 0, Optional caller As String, Optional Context As String, _
                            Optional Message As String, Optional newLine As eNewLine = Auto)
    If gcDebugMode Then
      outputLine mReport(d, caller, Context, Message, newLine)
    End If
End Sub
Public Sub forceReport(Optional d As Double = 0, Optional caller As String, Optional Context As String, _
                            Optional Message As String, Optional newLine As eNewLine = Auto)
    outputLine mReport(d, caller, Context, Message, newLine)
End Sub
Public Sub Log(line As String)
  If glogFile Is Nothing Then Set glogFile = New cTextStream
  glogFile.NextLine = line
End Sub
Private Sub outputLine(line As String)
  If gdebugOutoutToFile Then Log line
  Debug.Print line
End Sub

Private Function mReport(Optional d As Double = 0, Optional caller As String, Optional Context As String, _
                            Optional Message As String, Optional newLine As eNewLine = Auto) As String
Dim myContext As String
    
  If Len(caller) <> 0 Then mCaller = caller
  If Len(Context) = 0 Then
      If mFirstReport And Len(caller) > 0 Then
          myContext = "START"
'           Else
'               myContext = mContext
      End If
  Else
      mContext = Context
      myContext = mContext
  End If
  If newLine = Auto Then newLine = IIf(mFirstReport And gCallDepth = 1, Before, No)
  mReport = timeStamp(d:=d, newLine:=newLine, _
                          Indent:=gCallDepth, _
                          caller:=IIf(mFirstReport And Len(mCaller) > 0, mCaller, String(Len(mCaller), " ")), _
                          Context:=myContext, _
                          Message:=Message, _
                          dt:=IIf(mFirstReport Or Len(mCaller) = 0, "", FormatNumber(MicroTimer - lapTime, 6)))
  lapTime = MicroTimer

  mFirstReport = False
    
End Function
Property Let Context(RHS As String)
    mContext = RHS
End Property
Private Sub Class_Initialize()
    gCallDepth = gCallDepth + 1
    mFirstReport = True
    startTime = MicroTimer
    lapTime = startTime
End Sub
Private Sub Class_Terminate()
Dim line As String

    mContext = "END" & IIf(Len(ExitMessage) = 0, "", ": " & ExitMessage)
    If gcDebugMode Then
        outputLine timeStamp(newLine:=No, _
                                Indent:=gCallDepth, _
                                caller:=mCaller, _
                                Context:=mContext, _
                                dt:=FormatNumber(MicroTimer - startTime, 6))
        startTime = MicroTimer
   End If
    gCallDepth = gCallDepth - 1
    
End Sub

