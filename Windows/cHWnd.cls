VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' cpptWindow
' By Cool Blue cool.blue@y7mail.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' used by FormatMessage
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const FORMAT_MESSAGE_TEXT_LEN = &HA0


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows API Declares
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function FindWindow Lib "User32.dll" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "User32.dll" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long

Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    ByRef lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByRef Arguments As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" ( _
                    ByVal hWnd As Long) As Long
Private Declare Function getWindowText Lib "user32" Alias "GetWindowTextA" ( _
                    ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
 
Private Declare Function IsWindow Lib "user32" ( _
    ByVal hWnd As Long) As Long
 
Private Declare Function APIgetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hWnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long
    
'Get the process ID of this instance of Excel
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'Get the ID of the process that a window belongs to
Private Declare Function GetWindowThreadProcessId Lib "user32" ( _
                    ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
'
' Excel window class names
Private Const C_EXCEL_APPLICATION_CLASS = "XLMAIN"
Private Const C_EXCEL_DESKTOP_CLASS = "XLDESK"
Private Const C_EXCEL_WINDOW_CLASS = "EXCEL7"

Private mhWnd As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Handle(Wn) As Long
  Handle = FindHWndExFuzzy(Wn)
End Property
Property Get HandleByClass(Wn, ClassName As String)
  HandleByClass = FindHWndExFuzzy(W:=Wn, ClassName:=ClassName)
End Property
Property Get hCurrentProcess() As Long

End Property
Property Let hWnd(RHS As Long)
  mhWnd = RHS
End Property
Property Get Caption() As String
  Caption = getWindowTextHWnd(mhWnd)
End Property
Property Get ClassName() As String
  ClassName = getClassName(mhWnd)
End Property
Property Get ThreadPID() As Long
  If mhWnd <> 0 Then ThreadPID = GetWindowThreadProcessId(mhWnd, 0&)
End Property
'/////////////////////////////////////////////////////////////////////////////////////////////////////
'// Matches on captions that contain the supplied caption                                           //
'/////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function FindHWndExFuzzy(Optional W As Variant = vbNullString, _
                                Optional ClassName As Variant = vbNullString, _
                                Optional HWndSibling As Long = 0&, _
                                Optional HWndParent As Long = 0&, _
                                Optional NotCaption As String = vbNullString) As Long
Const myName As String = "cHWnd.FindHWndExFuzzy"
Dim db As New cDebugReporter

Dim Wn As Object, Caption As String
Dim WinHWnd As Long
Dim ErrNum As Long
Dim ErrText As String
Dim str As String
Dim isFound As Boolean

  db.Report caller:=myName, Context:="Class: " & vbTab & ClassName
                    
  
  Caption = getCaption(W)
  db.Report Message:="Looking for:" & vbTab & Caption & _
                      " excluding:" & vbTab & NotCaption
  
  '/////////////////////////////////////////////////////////////////////////////
  '// Have to use FindWindowEx to ensure all windows are found.               //
  '// cool.blue@y7mail.com                                                    //
  '/////////////////////////////////////////////////////////////////////////////
  
  WinHWnd = FindWindowEx(HWndParent, HWndSibling, ClassName, Caption)
  If WinHWnd = 0 Then
  '/////////////////////////////////////////////////////////////////////////////
  '// Windows piles all kinds of messages in the caption in square brackets.  //
  '// [Compatibility Mode], for example...                                    //
  '// Ignore this junk and search for the caption of the passed window.       //
  '// cool.blue@y7mail.com                                                    //
  '/////////////////////////////////////////////////////////////////////////////
    db.Report Message:="Exact match Not Found"
  ' Relax search to exclude caption...
    WinHWnd = FindWindowEx(HWndParent, HWndSibling, ClassName, vbNullString)
    Do While WinHWnd <> 0 And Not isFound
    ' Exit if target caption is included in found caption...
      str = getWindowTextHWnd(WinHWnd)
      db.Report Message:=str
      If InStr(str, Caption) > 0 And (NotCaption = vbNullString Or (InStr(str, NotCaption) = 0)) Then
          isFound = True
          db.Report Message:="Fuzzy match found: " & str & vbTab & Hex(WinHWnd)
          Exit Do
      Else
          WinHWnd = FindWindowEx(HWndParent, WinHWnd, ClassName, vbNullString)
      End If
    Loop
  End If

  If WinHWnd = 0 Then
      ''''''''''''''''''''''''''''''''
      ' an error occurred
      ''''''''''''''''''''''''''''''''
      'allow caller to check for 0 WinHWnd
          Exit Function
  End If
  
  db.ExitMessage = "WinHWnd: 0x" & Hex(WinHWnd)
  
  FindHWndExFuzzy = WinHWnd

End Function
Public Sub printWindows(Optional hRoot As Long = 0&)
Const myName As String = "cHWnd.printWindows"
Dim db As New cDebugReporter
    db.Report caller:=myName, Message:="Root: " & Hex(hRoot)
Dim fso As New FileSystemObject, f As TextStream
Dim hWnd As Long, ohWnd As cHWnd

  hWnd = FindWindowEx(hRoot, 0&, vbNullString, vbNullString)
  While hWnd
    Set ohWnd = New cHWnd
    With ohWnd
      .hWnd = hWnd
      db.forceReport Message:=Format(Hex(hWnd), "@@@@@@@@") & "," & .Caption & "," & .ClassName & "," & .ThreadPID
      hWnd = FindWindowEx(hRoot, hWnd, vbNullString, vbNullString)
    End With
  Wend
  
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////
'// Private functions                                                                               //
'/////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function getCaption(W As Variant) As String
  If W <> vbNullString Then
    getCaption = IIf(TypeOf W Is Object , W.Caption, _
                  IIf(TypeName(W) = "String", W, vbNullString))
  Else
    getCaption = vbNullString
  End If
End Function

Private Function TrimToNull(Text As String) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TrimToNull
' Returns the portion of Text that is to the left of
' the first vbNullChar character.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Pos As Integer
    Pos = InStr(1, Text, vbNullChar)
    If Pos Then
        TrimToNull = Left(Text, Pos - 1)
    Else
        TrimToNull = Text
    End If
End Function
Private Sub Class_Initialize()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class_Initialize
' This is called when the class is created.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub

Private Sub Class_Terminate()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class_Terminate
' This is called when the class terminates.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub


Private Function GetSystemErrorMessageText(ErrorNumber As Long) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetSystemErrorMessageText
'
' This function gets the system error message text that corresponds to the error code returned by the
' GetLastError API function or the Err.LastDllError property. It may be used ONLY for these error codes.
' These are NOT the error numbers returned by Err.Number (for these errors, use Err.Description to get
' the description of the message).
' The error number MUST be the value returned by GetLastError or Err.LastDLLError.
'
' In general, you should use Err.LastDllError rather than GetLastError because under some circumstances the value of
' GetLastError will be reset to 0 before the value is returned to VB. Err.LastDllError will always reliably return
' the last error number raised in a DLL.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim ErrorText As String
    Dim TextLen As Long
    Dim FormatMessageResult As Long
    Dim LangID As Long
    
    LangID = 0&
    ErrorText = String$(FORMAT_MESSAGE_TEXT_LEN, " ")
    TextLen = Len(ErrorText)
    On Error Resume Next
    FormatMessageResult = 0&
    
    FormatMessageResult = FormatMessage( _
                    dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                    lpSource:=0&, _
                    dwMessageId:=ErrorNumber, _
                    dwLanguageId:=0&, _
                    lpBuffer:=ErrorText, _
                    nSize:=TextLen, _
                    Arguments:=0&)
       
       
    On Error GoTo 0
    If FormatMessageResult > 0 Then
        ErrorText = TrimToNull(ErrorText)
        GetSystemErrorMessageText = ErrorText
    Else
        ' Format message didn't return any text. there is no text description for the specified error.
        GetSystemErrorMessageText = "NO ERROR DESCRIPTION AVAILABLE"
    End If
    
End Function
'/////////////////////////////////////////////////////////////
'//Extra routines to handle SDI and Compatibility Mode      //
'//cool.blue@y7mail.com                                     //
'/////////////////////////////////////////////////////////////
Private Function getWindowTextHWnd(hWnd As LongPtr) As String
Dim txt As String, txtLen As Long

    If hWnd = 0 Then
      getWindowTextHWnd = vbNullString
      Exit Function
    End If
    
    txtLen = GetWindowTextLength(hWnd) + 1
    txt = String(txtLen, vbNullChar)
    Call getWindowText(hWnd, txt, txtLen)
    getWindowTextHWnd = txt
    
End Function

Private Function getClassName(hWnd As LongPtr) As String
Dim strText As String, strLen As Long

    If hWnd = 0 Then
      getClassName = vbNullString
      Exit Function
    End If
    
    strText = String$(128, Chr$(0))
    strLen = APIgetClassName(hWnd, strText, 128)
    getClassName = Left(strText, strLen)
    
End Function
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


