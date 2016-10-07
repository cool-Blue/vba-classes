VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CScrollEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CScrollEventSink
' By Chip Pearson, chip@cpearson.com www.cpearson.com
'
' This procedure receives the events that are raised
' in the CSubClassActiveWindow class. The declaration
' Public WithEvents ScrollEvents As CSubClassActiveWindow
' defines an object named ScrollEvents of type
' CSubClassActiveWindow. The WithEvents keyword indicates
' that we're going to be receiving events from
' CSubClassActiveWindow.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declare WithEvents the class that is doing the subclassing
' and that calls RaiseEvent.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public WithEvents ScrollEvents As CDetectScroll
Attribute ScrollEvents.VB_VarHelpID = -1

' used by FormatMessage
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const FORMAT_MESSAGE_TEXT_LEN = &HA0


Private Declare Function PathCompactPathEx Lib "shlwapi.dll" Alias "PathCompactPathExA" ( _
    ByVal pszOut As String, _
    ByVal pszSrc As String, _
    ByVal cchMax As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function FindWindow Lib "User32.dll" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "User32.dll" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long

Private Declare Function GetClassName Lib "User32.dll" Alias "GetClassNameA" ( _
    ByVal HWnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Private Declare Function getWindowText Lib "User32.dll" Alias "GetWindowTextA" ( _
    ByVal HWnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long

Private Declare Function GetWindowTextLength Lib "User32.dll" Alias "GetWindowTextLengthA" ( _
    ByVal HWnd As Long) As Long

Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    ByRef lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByRef Arguments As Long) As Long

''''''''''''''''''''
' Used by SizeString
''''''''''''''''''''
Private Enum SizeStringSide
    TextLeft = 0
    TextRight = 1
End Enum



Private Sub ScrollEvents_ScrollEndHorizontalScroll(ByVal TopLeftCell As Range, ByVal Wnd As Window)
    HandleEvent "EndHorizontalScroll", TopLeftCell, Wnd
End Sub

Private Sub ScrollEvents_ScrollEndVerticalScroll(ByVal TopLeftCell As Range, ByVal Wnd As Window)
    HandleEvent "EndVerticalScroll", TopLeftCell, Wnd
End Sub

Private Sub ScrollEvents_ScrollLeft(ByVal TopLeftCell As Range, ByVal Wnd As Window)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ScrollEvents_ScrollLeft
' User scrolled left.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    HandleEvent "ScrollLeft", TopLeftCell, Wnd
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Event Procedures triggered by the RaiseEvent statements
' in CSubClassActiveWindow. These are called by RaiseEvent.
' The event procedure simply display text in the
' lbxMesssageBox listbox on form frMessageList.
' The function SizeString is defined in modSupport.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ScrollEvents_ScrollLineDown(ByVal TopLeftCell As Range, ByVal Wnd As Window)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ScrollEvents_ScrollLineDown
' User scrolled one line down.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    HandleEvent "ScrollLineDown", TopLeftCell, Wnd
End Sub

Private Sub ScrollEvents_ScrollLineLeft(ByVal TopLeftCell As Range, ByVal Wnd As Window)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ScrollEvents_ScrollLineLeft
' User scrolled line left.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    HandleEvent "ScrollLineLeft", TopLeftCell, Wnd
End Sub

Private Sub ScrollEvents_ScrollLineRight(ByVal TopLeftCell As Range, ByVal Wnd As Window)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ScrollEvents_ScrollLineRight
' User scrolled line right.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    HandleEvent "ScrollLineRight", TopLeftCell, Wnd
End Sub

Private Sub ScrollEvents_ScrollLineUp(ByVal TopLeftCell As Range, ByVal Wnd As Window)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ScrollEvents_ScrollLineUp
' User scrolled one line up.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    HandleEvent "ScrollLineUp", TopLeftCell, Wnd
End Sub

Private Sub ScrollEvents_ScrollPageDown(ByVal TopLeftCell As Range, ByVal Wnd As Window)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ScrollEvents_ScrollPageDown
' User scrolled one page down.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    HandleEvent "ScrollPageDown", TopLeftCell, Wnd
End Sub

Private Sub ScrollEvents_ScrollPageLeft(ByVal TopLeftCell As Range, ByVal Wnd As Window)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ScrollEvents_ScrollPageLeft
' User scrolled one page left.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    HandleEvent "ScrollPageLeft", TopLeftCell, Wnd
End Sub

Private Sub ScrollEvents_ScrollPageRight(ByVal TopLeftCell As Range, ByVal Wnd As Window)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ScrollEvents_ScrollPageRight
' User scrolled one page right.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    HandleEvent "ScrollPageRight", TopLeftCell, Wnd
End Sub

Private Sub ScrollEvents_ScrollPageUp(ByVal TopLeftCell As Range, ByVal Wnd As Window)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ScrollEvents_ScrollPageUp
' User scrolled one page up.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    HandleEvent "ScrollPageUp", TopLeftCell, Wnd
End Sub

Private Sub ScrollEvents_ScrollRight(ByVal TopLeftCell As Range, ByVal Wnd As Window)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ScrollEvents_ScrollRight
' User scrolled right.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    HandleEvent "ScrollRight", TopLeftCell, Wnd
End Sub



Private Sub HandleEvent(EventName As String, TopLeftCell As Range, Wnd As Window)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HandleEvent
' All the events do the same thing -- adding text to the listbox.
' This procedure does the same thing for all events.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim S As String
    
    S = CreateMessageString(EventName, TopLeftCell, Wnd)

    frmMessageList.Show vbModeless
    frmMessageList.lbxMessageBox.AddItem S
    frmMessageList.lblActiveWindowText.Caption = "Active Window: " & Application.ActiveWindow.Caption
    With frmMessageList.lbxMessageBox
        .TopIndex = .ListCount - 1
    End With

End Sub


Private Function CreateMessageString(EventName As String, TopLeftCell As Range, Wnd As Window) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CreateMessageString
' Creates the message string that is added to the listbox.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim S As String
    S = SizeString(EventName, 20, TextLeft, " ") & _
        SizeString("TopLeftCell: " & TopLeftCell.Address(False, False), 20, TextLeft, " ") & _
        SizeString("Window Caption: " & Wnd.Caption, 40)
        
        

    CreateMessageString = S

End Function

Private Function SizeString(Text As String, Length As Long, _
    Optional ByVal TextSide As SizeStringSide = TextLeft, _
    Optional PadChar As String = " ") As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SizeString
' This procedure creates a string of a specified length. Text is the original string
' to include, and Length is the length of the result string. TextSide indicates whether
' Text should appear on the left (in which case the result is padded on the right with
' PadChar) or on the right (in which case the string is padded on the left). When padding on
' either the left or right, padding is done using the PadChar. character. If PadChar is omitted,
' a space is used. If PadChar is longer than one character, the left-most character of PadChar
' is used. If PadChar is an empty string, a space is used. If TextSide is neither
' TextLeft or TextRight, the procedure uses TextLeft.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim sPadChar As String

If Len(Text) >= Length Then
    ' if the source string is longer than the specified length, return the
    ' Length left characters
    SizeString = Left(Text, Length)
    Exit Function
End If

If Len(PadChar) = 0 Then
    ' PadChar is an empty string. use a space.
    sPadChar = " "
Else
    ' use only the first character of PadChar
    sPadChar = Left(PadChar, 1)
End If

If (TextSide <> TextLeft) And (TextSide <> TextRight) Then
    ' if TextSide was neither TextLeft nor TextRight, use TextLeft.
    TextSide = TextLeft
End If

If TextSide = TextLeft Then
    ' if the text goes on the left, fill out the right with spaces
    SizeString = Text & String(Length - Len(Text), sPadChar)
Else
    ' otherwise fill on the left and put the Text on the right
    SizeString = String(Length - Len(Text), sPadChar) & Text
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

Private Function GetWindowClassName(HWnd As Long) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetWindowClassName
' Returns the window class name of HWnd.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Res As Long
    Dim className As String
    Dim ErrNum As Long
    Dim ErrText As String
    
    className = String$(255, vbNullChar)
    Res = GetClassName(HWnd, className, 255)
    If Res = 0 Then
        ''''''''''''''''''''''''
        ' An error occurred
        ''''''''''''''''''''''''
        ErrNum = Err.LastDllError
        ErrText = GetSystemErrorMessageText(ErrNum)
        MsgBox "Error On GetClassName" & vbCrLf & _
            "Err: " & CStr(ErrNum) & vbCrLf & _
            "Description: " & ErrText
        GetWindowClassName = vbNullString
        Exit Function
    End If
    
    GetWindowClassName = TrimToNull(className)

End Function

Private Function GetWindowTextString(HWnd As Long) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetWindowTextString
' Returns the window text of HWnd
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Res As Long
    Dim WindowText As String
    Dim WindowTextLen As Long
    Dim ErrNum As Long
    Dim ErrText As String
        
    Res = GetWindowTextLength(HWnd)
    If Res = 0 Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Res = 0 indicates 0 length window text or an error.
        ' Test Err.LastDllErrror to see if we had an error.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Err.LastDllError <> 0 Then
            '''''''''''''''''''''''''''''''''''''''''''
            ' An error occurred
            '''''''''''''''''''''''''''''''''''''''''''
            ErrNum = Err.LastDllError
            ErrText = GetSystemErrorMessageText(ErrNum)
            MsgBox "Error With GetWindowTextLength" & vbCrLf & _
                "Err: " & CStr(ErrNum) & vbCrLf & _
                "Desc: " & ErrText
            GetWindowTextString = vbNullString
            Exit Function
        End If
    Else
        '''''''''''''''''
        ' Res <> 0
        '''''''''''''''''
        WindowText = String$(Res + 1, vbNullChar)
        Res = getWindowText(HWnd, WindowText, Res + 1)
        If Res = 0 Then
            ''''''''''''''''''''''''''''''''''''''''''
            ' We had either a 0 lenght Window Text
            ' or an error. Test Err.LastDllError
            ' to see if we had an error.
            ''''''''''''''''''''''''''''''''''''''''''
            If Err.LastDllError <> 0 Then
                ''''''''''''''''''''''''''''''
                ' An error occurred.
                '''''''''''''''''''''''''''''
                ErrNum = Err.LastDllError
                ErrText = GetSystemErrorMessageText(ErrNum)
                MsgBox "Error With GetWindowText" & vbCrLf & _
                    "Err: " & CStr(ErrNum) & vbCrLf & _
                    "Desc: " & ErrText
                GetWindowTextString = vbNullString
            End If
        End If
    End If
    GetWindowTextString = TrimToNull(Text:=WindowText)
End Function


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
Public Function ShortenTextToChars(InputText As String, _
        NumberOfCharacters As Long) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShortenTextToChars
' This function returns a shortened version of the InputText parameter that is
' NumberOfCharacters in length. This function is primarily designed for use with
' fully-qualified file names. With a file name, the function will remove or truncate
' an element of the path (e.g., a folder name), replacing the removed text with the string
' "...". While this is intended for use with file names, it will work with any text string.
' When used on text that does not contain '\' characters, it typically just truncates the
' right side of InputText.
' Returns vbNullString is an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim ResString As String
Dim Res As Long
Dim ErrorNumber As Long
Dim ErrorText As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure that InputText is not an empty string
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If InputText = vbNullString Then
    MsgBox "The InputText parameter is an empty string"
    ShortenTextToChars = vbNullString
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Test for NumberOfCharacters <= 3. If the InputText
' is 3 or fewer characters, PathCompactPathEx would replace the
' entire string with "...". We don't want that. Return the entire
' InputText.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Len(InputText) <= 3 Then
    ShortenTextToChars = InputText
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Test for NumberOfCharacters less than equal to 3.
' If the NumberOfCharacters <= 3, PathCompactPathEx would replace
' the entire InputString with "...".
' Instead, return the left-most characters and get out.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If NumberOfCharacters <= 3 Then
    ShortenTextToChars = Left(InputText, NumberOfCharacters)
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a valid number of characters. If NumberOfCharacters
' is less than or equal to 0, or greater than the length of
' the InputText, PathCompactPathEx will crash the application.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If NumberOfCharacters <= 0 Then
    MsgBox "The NumberOfCharacters must be greater than 0."
    ShortenTextToChars = vbNullString
    Exit Function
End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Here we test if the length of InputTex is equal to
' NumberOfCharacters. If they are equal, return the
' the entire string and exit. If we allowed
' PathCompactPathEx to process the string, it would truncate
' on the right and replace the last three characters on the
' right with "...". We don't want that behavior -- we
' want to return the entire string.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Len(InputText) = NumberOfCharacters Then
    ShortenTextToChars = InputText
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialize the buffer. When PathCompactPathEx
' creates its string, it considers NumberOfCharacters
' to include room for the trailing null character. Thus
' the actual number of real characters it returns will be
' NumberOfCharacters-1. Thus, we allocate the string
' to NumberOfCharacters+2 = 1 because we want
' NumberOfCharacters (without the trailing null)
' returned, + 1 for trailing null.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
ResString = String$(NumberOfCharacters + 2, vbNullChar)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shorten the text with PathCompactPathEx
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Res = PathCompactPathEx(ResString, InputText, NumberOfCharacters, 0&)
If Res = 0 Then
    ErrorNumber = Err.LastDllError
    ErrorText = GetSystemErrorMessageText(ErrorNumber)
    MsgBox "An error occurred with PathCompactPathEx" & vbCrLf & _
           "Error Number: " & CStr(ErrorNumber) & vbCrLf & _
           "Descrition: " & ErrorText
    ShortenTextToChars = vbNullString
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''
' trim to get the charcters to the left
' of the vbNullChar character.
'''''''''''''''''''''''''''''''''''''''
ResString = TrimToNull(Text:=ResString)

'''''''''''''''''''''''''''''''''''''''
' return the result string
'''''''''''''''''''''''''''''''''''''''
ShortenTextToChars = ResString

End Function

