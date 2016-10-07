VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSystemTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Private mZeroOffset As SYSTEMTIME

Public Function sysTime() As String
Attribute sysTime.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute sysTime.VB_UserMemId = 0
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
    sysTime = toString(MyTime)
exitEarly:
End Function
Public Function offsetSysTime(secOffs As Double, Optional rawTime As String) As String
Dim MyTime As SYSTEMTIME, tOffs As SYSTEMTIME

    GetLocalTime MyTime
    With tOffs
      .wSecond = Int(secOffs)
      .wMilliseconds = Int((secOffs - .wSecond) * 1000)
    End With
    On Error Resume Next
    rawTime = toString(MyTime)
    MyTime = sysTimeAdd(MyTime, tOffs)
    offsetSysTime = toString(MyTime)
    
exitEarly:
End Function
Private Function toString(MyTime As SYSTEMTIME) As String
    toString = Format(MyTime.wHour, "00") & ":" & Format(MyTime.wMinute, "00") & ":" & _
                Format(MyTime.wSecond, "00") & ":" & Format(MyTime.wMilliseconds, "000")
End Function
Private Function sysTimeAdd(t As SYSTEMTIME, dt As SYSTEMTIME) As SYSTEMTIME
Dim MyTime As SYSTEMTIME, carry As Integer, tDatePart As Long
  tDatePart = DateSerial(t.wYear, t.wMonth, t.wDay)
  With MyTime
    addWithCarry t.wMilliseconds, dt.wMilliseconds, 1000, .wMilliseconds, carry
    addWithCarry t.wSecond + carry, dt.wSecond, 60, .wSecond, carry
    addWithCarry t.wMinute + carry, dt.wMinute, 60, .wMinute, carry
    addWithCarry t.wHour + carry, dt.wHour, 24, .wHour, carry
    tDatePart = tDatePart + carry + DateSerial(dt.wYear, dt.wMonth, dt.wDay)
    .wDay = Day(tDatePart)
    .wMonth = Month(tDatePart)
    .wYear = Year(tDatePart)
  End With
  sysTimeAdd = MyTime
End Function
Private Sub addWithCarry(n1 As Integer, n2 As Integer, modulus As Integer, n As Integer, c As Integer)
  
  n = (n1 + n2) Mod modulus
  c = (n1 + n2 - n) / modulus
  
End Sub
