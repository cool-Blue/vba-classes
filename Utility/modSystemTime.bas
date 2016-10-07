Attribute VB_Name = "modSystemTime"
Option Explicit

Const cModuleIndent = 1
Public Enum eNewLine
    No
    Before
    After
    Both
End Enum
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
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
    With Application.WorksheetFunction
        sysTime = .Text(MyTime.wHour, "00") & ":" & .Text(MyTime.wMinute, "00") & ":" & _
                    .Text(MyTime.wSecond, "00") & ":" & .Text(MyTime.wMilliseconds, "0000")
    End With
End Function
Public Function timeStamp(Optional d As Double = 0, Optional newLine As eNewLine = No, Optional Indent As Long = 0, _
                            Optional caller As String, Optional context As String, Optional contextCol As Long = -1, _
                            Optional message As String, Optional messageCol As Long = -1) As String
Dim errorMessage As String

    If Err.Number <> 0 Then
        errorMessage = "ERROR: " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    If d = 0 Then
        timeStamp = sysTime
    Else
        With Application.WorksheetFunction
            timeStamp = .Text(Hour(d), "00") & ":" & .Text(Minute(d), "00") & ":" & .Text(Second(d), "00") & ":" & rept(Chr(9), Indent)
        End With
    End If
    If Len(caller) <> 0 Then timeStamp = timeStamp & vbTab & caller
    If Len(context) <> 0 Then
        timeStamp = timeStamp & setCol(timeStamp, context, contextCol)
    End If
    If Len(message) <> 0 Then
        timeStamp = timeStamp & setCol(timeStamp, message, messageCol)
    End If
    If Len(errorMessage) <> 0 Then
        If Len(timeStamp) <> 0 Then timeStamp = timeStamp & vbCrLf
        timeStamp = timeStamp & Chr(9) & errorMessage
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

Function setCol(s1 As String, s2 As String, s2Col As Long, Optional divider As String = ":") As String
    If s2Col < 0 Then
        setCol = Chr(9) & divider & s2
    Else
        setCol = rept(" ", max(s2Col - Len(s1) - Len(divider), 0)) & divider & s2
    End If
End Function


