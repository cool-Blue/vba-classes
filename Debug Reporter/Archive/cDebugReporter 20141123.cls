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
Public exitMessage As String

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
                          ByVal hwnd As Long, _
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
                          ByVal hwnd As Long, ByVal wMsg As Long, _
                          ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                          ByVal hwnd As Long, ByVal wMsg As Long, _
                          ByVal wParam As Long, ByVal lParam As Long) As Long


Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function getWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
 
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_SETFOCUS As Long = &H7

Private Const KEYSTATE_KEYDOWN As Long = &H80
'Private savState(0 To 255) As Byte

Private Const TABSTOP As String = " "

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
Public Function fmtTime(t As Double) As String
  fmtTime = Format(Hour(t), "00") & ":" & Format(Minute(t), "00") & ":" & Format(Second(t), "00")
End Function
Public Function timeStamp(Optional d As Double = 0, Optional newLine As eNewLine = No, Optional Indent As Long = 0, _
                            Optional Caller As String, Optional Context As String, Optional message As String, _
                            Optional dt As String, Optional dtCol As Long = 120) As String
Dim errorMessage As String, hasCaller As Boolean

    If Err.Number <> 0 Then
        errorMessage = "ERROR: " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    
    If d = 0 Then
        timeStamp = sysTime & String(Indent, TABSTOP)
    Else
      timeStamp = fmtTime(d) & ":" & String(Indent, TABSTOP)
    End If
    
    If Len(Caller) <> 0 Then
      hasCaller = True
      timeStamp = timeStamp & Caller
      
      If Len(Context) <> 0 Then timeStamp = timeStamp & ": " & TABSTOP & Context
      If Len(message) <> 0 Then timeStamp = timeStamp & IIf(Len(Context) = 0, ": ", TABSTOP) & message
      
    Else
    
      If Len(Context) <> 0 Then timeStamp = timeStamp & TABSTOP & Context
      If Len(message) <> 0 Then timeStamp = timeStamp & IIf(Len(Context) = 0, TABSTOP, "  ") & message
      
    End If
        
    If Len(dt) <> 0 Then
      If Len(errorMessage) <> 0 Then
        errorMessage = sysTime & String(Indent, "*") & errorMessage
        errorMessage = errorMessage & setCol(errorMessage, dt, dtCol)
        timeStamp = timeStamp & vbNewLine & errorMessage
      Else
        timeStamp = timeStamp & setCol(timeStamp, dt, dtCol)
      End If
    Else
      If Len(errorMessage) <> 0 Then
        errorMessage = sysTime & String(Indent, "*") & errorMessage
        timeStamp = timeStamp & vbNewLine & errorMessage
      End If
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

Public Sub Report(Optional message As String, Optional d As Double = 0, Optional Caller As String, _
                            Optional Context As String, Optional newLine As eNewLine = Auto)
Attribute Report.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Report.VB_UserMemId = 0
If gcDebugMode Then
      outputLine mReport(d, Caller, Context, message, newLine)
    End If
End Sub
Public Sub forceReport(Optional d As Double = 0, Optional Caller As String, Optional Context As String, _
                            Optional message As String, Optional newLine As eNewLine = Auto)
    outputLine mReport(d, Caller, Context, message, newLine)
End Sub
Private Sub Log(line As String)
  If glogFile Is Nothing Then Set glogFile = New cTextStream
  glogFile.NextLine = line
End Sub
Private Sub outputLine(line As String)
  If gdebugOutoutToFile Then Log line
  Debug.Print line
End Sub

Private Function mReport(Optional d As Double = 0, Optional Caller As String, Optional Context As String, _
                            Optional message As String, Optional newLine As eNewLine = Auto) As String
Dim myContext As String
    
  If Len(Caller) <> 0 Then mCaller = Caller
  If Len(Context) = 0 Then
      If mFirstReport And Len(Caller) > 0 Then
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
                          Caller:=Caller, _
                          Context:=myContext, _
                          message:=message, _
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

    If Not mFirstReport Then
      mContext = "END" & IIf(Len(mContext) = 0, "", ": " & mContext) & _
      IIf(Len(exitMessage) = 0, "", ": " & exitMessage)
    Else
      mContext = mContext & _
      IIf(Len(exitMessage) = 0, "", ": " & exitMessage)
    End If
    If gcDebugMode Then
        outputLine timeStamp(newLine:=No, _
                                Indent:=gCallDepth, _
                                Caller:=mCaller, _
                                Context:=mContext, _
                                dt:=FormatNumber(MicroTimer - startTime, 6))
        startTime = MicroTimer
   End If
    gCallDepth = gCallDepth - 1
    
End Sub

