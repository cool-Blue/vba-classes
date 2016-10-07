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
'/ Public gCallDepth as Long
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
    With Application.WorksheetFunction
        On Error Resume Next
        sysTime = .Text(MyTime.wHour, "00") & ":" & .Text(MyTime.wMinute, "00") & ":" & _
                    .Text(MyTime.wSecond, "00") & ":" & .Text(MyTime.wMilliseconds, "000")
    End With
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
        With Application.WorksheetFunction
            timeStamp = .Text(Hour(d), "00") & ":" & .Text(Minute(d), "00") & ":" & .Text(Second(d), "00") & ":" & String(Indent, " ")
        End With
    End If
    If Len(caller) <> 0 Then timeStamp = timeStamp & " " & caller
    If Len(Context) <> 0 Then timeStamp = timeStamp & ": " & " " & Context
    If Len(Message) <> 0 Then timeStamp = timeStamp & ": " & " " & Message
    If Len(errorMessage) <> 0 Then
        timeStamp = timeStamp & " " & errorMessage
    End If
    If Len(dt) <> 0 Then
        timeStamp = timeStamp & setCol(timeStamp, dt, dtCol)
    End If
    Select Case newLine
    Case Before
        timeStamp = Chr(10) & timeStamp
    Case After
        timeStamp = timeStamp & Chr(10)
    Case Both
        timeStamp = Chr(10) & timeStamp & Chr(10)
    Case Else
    End Select
    
End Function
Function setCol(S1 As String, S2 As String, s2Col As Long, Optional divider As String = ":") As String
    
    If s2Col < 0 Then
        setCol = " " & divider & S2
    Else
        setCol = String(Application.max(s2Col - Len(S1) - Len(divider), 0), " ") & divider & S2
    End If
    
End Function

Private Sub Class_Initialize()
    gCallDepth = gCallDepth + 1
    mFirstReport = True
    startTime = MicroTimer
    lapTime = startTime
End Sub
Public Sub Report(Optional d As Double = 0, Optional caller As String, Optional Context As String, _
                            Optional Message As String, Optional newLine As eNewLine = Auto)

    If gcDebugMode Then
    Dim myContext As String
        
        If Len(caller) <> 0 Then mCaller = caller
        If Len(Context) = 0 Then
            If mFirstReport And Len(caller) > 0 Then
                myContext = "START"
            Else
                myContext = mContext
            End If
        Else
            mContext = Context
            myContext = mContext
        End If
        If newLine = Auto Then newLine = IIf(mFirstReport And gCallDepth = 1, Before, No)
        Debug.Print timeStamp(d:=d, newLine:=newLine, _
                                Indent:=gCallDepth, _
                                caller:=IIf(mFirstReport And Len(mCaller) > 0, mCaller, String(Len(mCaller), " ")), _
                                Context:=myContext, _
                                Message:=Message, _
                                dt:=IIf(mFirstReport Or Len(mCaller) = 0, "", Text(MicroTimer - lapTime, 6)))
        lapTime = MicroTimer

        mFirstReport = False
    End If
    
End Sub
Property Let Context(RHS As String)
    mContext = RHS
End Property
Private Sub Class_Terminate()
    mContext = "END" & IIf(Len(ExitMessage) = 0, "", ": " & ExitMessage)
    If gcDebugMode Then
        Debug.Print timeStamp(newLine:=No, _
                                Indent:=gCallDepth, _
                                caller:=mCaller, _
                                Context:=mContext, _
                                dt:=Text(MicroTimer - startTime, 6))
        startTime = MicroTimer
   End If
    gCallDepth = gCallDepth - 1
End Sub

