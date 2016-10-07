VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDetectScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CDetectScroll
' By Chip Pearson, www.cpearson.com chip@cpearson.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' The Excel object model does not provide an event when the user scrolls a worksheet window. If you need
' to detect when a window is scrolled, for example to reposition some controls on a worksheet, the closest
' you can get, using only the built-in methods, is using the SelectionChange event or periodically testing
' the visible range of a window.
'
' This class is used to detect when the user scrolls a worksheet window. It does this by subclassing
' the active window and listening for  WM_VSCROLL (vertical scroll) and WM_HSCROLL (horizontal scroll)
' messages. To detect scrolling in your workbook, you need include ONLY this class. The module SSubTmr6,
' the form frmMessageList, and the class CScrollEventSink are for demonstration only. This class
' is entirely self contained.
'
' This class REQUIRES the SSubTmr6.dll available for free at http://www.vbaccelerator.com/codelib/ssubtmr/ssubtmr.htm.
' See the code module modSSubTmr6 for documentation. There is no code in that module -- it is documentation only.
' See http://www.vbaccelerator.com/home/VB/Code/Libraries/Subclassing/SSubTimer/article.asp for documentation
' about SSubTmr6.dll.

' See also www.cpearson.com/excel/SubclassingWithSSubtmr6.htm for additional documentation about
' the SSubTmr6.dll.
'
' This class raises the following events:
'            ScrollLineDown
'            ScrollLineUp
'            ScrollPageDown
'            ScrollPageUp
'            ScrollLeft
'            ScrollRight
'            ScrollLineLeft
'            ScrollLineRight
'            ScrollPageLeft
'            ScrollPageRight
'            ScrollEndScroll
'
' Each of these events passes in a reference to the TopLeftCell and a reference to the window being scrolled.
'
' ----------------------------------------------------------------------------------------------------------
' Usage:
' ----------------------------------------------------------------------------------------------------------
' You need to determine what object module you want to receive
' the event messages. This can be a separate class module (as it is in
' this demonstration workbook), in a Userform's code module, in the
' ThisWorkbook module, or any sheet's code module.
' Regardless of where you decide to declare the object and receive the
' events, use code like the following:
'       Dim WithEvents ScrollEvents As CDetectScroll
'       Sub InitializeDetectScroll()
'           Set ScrollEvents = New CDetectScroll
'       End Sub
'
' Then declare the event procedures. For example,
'
'        Private Sub ScrollEvents_ScrollPageDown(ByVal TopLeftCell As Range, ByVal Wnd As Window)
'            Debug.Print "ScrollPageDown", TopLeftCell.Address, Wnd.Caption
'        End Sub
'
' To enabled scroll detection in your workbook, you need include ONLY this class module.
'
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' NOTE: SSubTmr6.dll can crash Excel when you are
' presently subclassing an Excel window AND you are
' editing the code in the VBA Editor.
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal HWnd As Long) As Long
Private Declare Function getWindowText Lib "user32" Alias "GetWindowTextA" (ByVal HWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
 
Private Declare Function IsWindow Lib "user32" ( _
    ByVal HWnd As Long) As Long
 
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal HWnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long
     

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Event Declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Event ScrollLineDown(ByVal TopLeftCell As Range, ByVal Wnd As Excel.Window)
Public Event ScrollLineUp(ByVal TopLeftCell As Range, ByVal Wnd As Excel.Window)
Public Event ScrollPageDown(ByVal TopLeftCell As Range, ByVal Wnd As Excel.Window)
Public Event ScrollPageUp(ByVal TopLeftCell As Range, ByVal Wnd As Excel.Window)
Public Event ScrollLeft(ByVal TopLeftCell As Range, ByVal Wnd As Excel.Window)
Public Event ScrollRight(ByVal TopLeftCell As Range, ByVal Wnd As Excel.Window)
Public Event ScrollLineLeft(ByVal TopLeftCell As Range, ByVal Wnd As Excel.Window)
Public Event ScrollLineRight(ByVal TopLeftCell As Range, ByVal Wnd As Excel.Window)
Public Event ScrollPageLeft(ByVal TopLeftCell As Range, ByVal Wnd As Excel.Window)
Public Event ScrollPageRight(ByVal TopLeftCell As Range, ByVal Wnd As Excel.Window)
Public Event ScrollEndHorizontalScroll(ByVal TopLeftCell As Range, ByVal Wnd As Excel.Window)
Public Event ScrollEndVerticalScroll(ByVal TopLeftCell As Range, ByVal Wnd As Excel.Window)
''''''''''''''''''''''''''''''''''''''''''''
' Implement the ISubclass interface. This
' is defined in SSubTimer6.
''''''''''''''''''''''''''''''''''''''''''''
Implements SSubTimer6.ISubclass

''''''''''''''''''''''''''''''''''''''''
' Private Constants
''''''''''''''''''''''''''''''''''''''''
Private Const WM_VSCROLL As Long = &H115
Private Const WM_HSCROLL As Long = &H114

Private Const SB_BOTH As Long = 3
Private Const SB_ENDSCROLL As Long = 8
Private Const SB_THUMBPOSITION As Long = 4
Private Const SB_THUMBTRACK As Long = 5
Private Const SB_TOP As Long = 6


Private Const SB_LINEUP As Long = 0
Private Const SB_LINEDOWN As Long = 1

Private Const SB_PAGEUP As Long = 2
Private Const SB_PAGEDOWN As Long = 3

Private Const SB_RIGHT As Long = 7
Private Const SB_LEFT As Long = 6

Private Const SB_LINELEFT As Long = 0
Private Const SB_PAGELEFT As Long = 2

Private Const SB_LINERIGHT As Long = 1
Private Const SB_PAGERIGHT As Long = 3

' Excel window class names
Private Const C_EXCEL_APPLICATION_CLASS = "XLMAIN"
Private Const C_EXCEL_DESKTOP_CLASS = "XLDESK"
Private Const C_EXCEL_WINDOW_CLASS = "EXCEL7"

' used by FormatMessage
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const FORMAT_MESSAGE_TEXT_LEN = &HA0

'''''''''''''''''''''''''''''''''''''''''''''
' We need to save the HWnd so we can call
' DetachMessage to stop subclassing on that
' window and optionally start subclassing
' a new window. We use the Workbook_WindowActivate
' in ThisWorkbook to detect when a window
' is activated.
''''''''''''''''''''''''''''''''''''''''''''''
Private PrevHWnd As Long
Private pEnabled As Boolean


''''''''''''''''''''''''''''''''''''''''''''''
' Excel Application variable
''''''''''''''''''''''''''''''''''''''''''''''
Private WithEvents XL As Excel.Application
Attribute XL.VB_VarHelpID = -1

''''''''''''''''''''''''''''''''''''''''''''''
' Enabled Property.
' This property allows the user's code to
' enable or disable scroll events.
''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Enabled() As Boolean
    Enabled = pEnabled
End Property
Public Property Let Enabled(Value As Boolean)
    pEnabled = Value
End Property


Private Sub Class_Initialize()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class_Initialize
' This is called when the class is created.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim HWnd As Long

Set XL = Application
pEnabled = True

If Not XL.ActiveWindow Is Nothing Then
    If XL.ActiveWindow.Type = xlWorkbook Then
        HWnd = GetHWndOfWindow(Wn:=Application.ActiveWindow)
        Debug.Print "HWnd: " & String(8 - Len(Hex(HWnd)), 0) & Hex(HWnd)
        If HWnd > 0 Then
            StartSubClassing (HWnd)
        End If
    End If
End If

End Sub

Private Sub Class_Terminate()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class_Terminate
' This is called when the class terminates.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
StopSubClassing

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implemented Procedures
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
' we don't use this property, but it is required by Implements
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Property

Private Function ISubclass_WindowProc(ByVal HWnd As Long, ByVal iMsg As Long, ByVal WParam As Long, ByVal LParam As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ISubclass_WindowProc
' This procedure is called automatically by SSubTimer6 when the use scrolls. The WParam variable
' indicates how the user scrolled (up/down, left/right, page/line. LParam is the handle
' of the scroll bar used to scroll. We don't use LParam.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''
' If the Enabled property is False, exit
' the function.
''''''''''''''''''''''''''''''''''''''''''
If pEnabled = False Then
    Exit Function
End If

Select Case iMsg
    Case WM_VSCROLL
    '''''''''''''''''''''''''''''''''''''''''''
    ' The user scrolled vertically
    '''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''
        ' WParam indicates how much
        ' and in what direction we
        ' scrolled (line/page, up/down).
        '''''''''''''''''''''''''''''''''
        Select Case WParam
            Case SB_LINEDOWN
                RaiseEvent ScrollLineDown(ActiveWindow.VisibleRange.Cells(1, 1), Application.ActiveWindow)
            Case SB_LINEUP
                RaiseEvent ScrollLineUp(ActiveWindow.VisibleRange.Cells(1, 1), Application.ActiveWindow)
            Case SB_PAGEUP
                RaiseEvent ScrollPageUp(ActiveWindow.VisibleRange.Cells(1, 1), Application.ActiveWindow)
            Case SB_PAGEDOWN
                RaiseEvent ScrollPageDown(ActiveWindow.VisibleRange.Cells(1, 1), Application.ActiveWindow)
            Case SB_ENDSCROLL
                RaiseEvent ScrollEndVerticalScroll(ActiveWindow.VisibleRange.Cells(1, 1), Application.ActiveWindow)
            Case Else
        End Select
    
    Case WM_HSCROLL
    ''''''''''''''''''''''''''''''''''''''''''''
    ' The user scrolled horizontally
    ''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''
        ' WParam indicates how much
        ' and in what direction we
        ' scrolled (line/page, left/right).
        '''''''''''''''''''''''''''''''''
        Select Case WParam
            Case SB_LEFT
                RaiseEvent ScrollLeft(ActiveWindow.VisibleRange.Cells(1, 1), Application.ActiveWindow)
            Case SB_RIGHT
                RaiseEvent ScrollRight(ActiveWindow.VisibleRange.Cells(1, 1), Application.ActiveWindow)
            Case SB_LINELEFT
                RaiseEvent ScrollLineLeft(ActiveWindow.VisibleRange.Cells(1, 1), Application.ActiveWindow)
            Case SB_LINERIGHT
                RaiseEvent ScrollLineRight(ActiveWindow.VisibleRange.Cells(1, 1), Application.ActiveWindow)
            Case SB_PAGELEFT
                RaiseEvent ScrollPageLeft(ActiveWindow.VisibleRange.Cells(1, 1), Application.ActiveWindow)
            Case SB_PAGERIGHT
                RaiseEvent ScrollPageRight(ActiveWindow.VisibleRange.Cells(1, 1), Application.ActiveWindow)
            Case SB_ENDSCROLL
                RaiseEvent ScrollEndHorizontalScroll(ActiveWindow.VisibleRange.Cells(1, 1), Application.ActiveWindow)
            Case Else
        End Select
    
    Case Else
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Some other message. We should never get here.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Select

End Function

Private Sub StartSubClassing(HWnd As Long)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' StartSubclassing
' This starts subclassing. If first calls DetachMessage
' to stop subclassing the previous window and start
' subclassing the current window.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    SSubTimer6.DetachMessage Me, PrevHWnd, WM_HSCROLL
    SSubTimer6.DetachMessage Me, PrevHWnd, WM_VSCROLL
    On Error GoTo 0
    
    SSubTimer6.AttachMessage Me, HWnd, WM_HSCROLL
    SSubTimer6.AttachMessage Me, HWnd, WM_VSCROLL
    PrevHWnd = HWnd

End Sub

Private Sub StopSubClassing()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' StopSubClassing
' This calls DetachMessage to stop subclassing.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    SSubTimer6.DetachMessage Me, PrevHWnd, WM_HSCROLL
    SSubTimer6.DetachMessage Me, PrevHWnd, WM_VSCROLL
    On Error GoTo 0

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Excel Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub XL_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    Dim HWnd As Long
    If Wn.Type = xlWorkbook Then
        HWnd = GetHWndOfWindow(Wn)
        If HWnd <> 0 Then
            StartSubClassing HWnd:=HWnd
        End If
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private Procedures
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetHWndOfWindow(Wn As Excel.Window)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetHWndOfWindow
' This returns the window handle (HWnd) of the window referenced by Wn.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim XLHWnd As Long
Dim DeskHWnd As Long
Dim WinHWnd As Long
Dim ErrNum As Long
Dim ErrText As String
Dim str As String
Dim isFound As Boolean

Debug.Print "Looking for:", Wn.Caption

'///////////////////////////////////////////////////////////////
'//Have to use FindWindowEx to ensure all windows are found.  //
'//cool.blue@y7mail.com                                       //
'///////////////////////////////////////////////////////////////
XLHWnd = FindWindowEx(0&, 0&, C_EXCEL_APPLICATION_CLASS, vbNullString)

'Loop in case of SDI (XL 2013)
Do While XLHWnd <> 0 And Not isFound
    If XLHWnd = 0 Then
        ''''''''''''''''''''''''''''''''
        ' an error occurred
        ''''''''''''''''''''''''''''''''
        ErrNum = Err.LastDllError
        ErrText = GetSystemErrorMessageText(ErrorNumber:=ErrNum)
        MsgBox "An error occurred with FindWindow for " & C_EXCEL_APPLICATION_CLASS & vbCrLf & _
            "Error: " & CStr(ErrNum) & "  Hex(" & Hex(ErrNum) & ")" & vbCrLf & _
            "Description: " & ErrText
            GetHWndOfWindow = 0
            Exit Function
    End If
    
    DeskHWnd = FindWindowEx(XLHWnd, 0&, C_EXCEL_DESKTOP_CLASS, vbNullString)
    If DeskHWnd = 0 Then
        ''''''''''''''''''''''''''''''''
        ' an error occurred
        ''''''''''''''''''''''''''''''''
        ErrNum = Err.LastDllError
        ErrText = GetSystemErrorMessageText(ErrorNumber:=ErrNum)
        MsgBox "An error occurred with FindWindow for " & C_EXCEL_DESKTOP_CLASS & vbCrLf & _
            "Error: " & CStr(ErrNum) & "  Hex(" & Hex(ErrNum) & ")" & vbCrLf & _
            "Description: " & ErrText
            GetHWndOfWindow = 0
            Exit Function
    End If
    
    WinHWnd = FindWindowEx(DeskHWnd, 0&, C_EXCEL_WINDOW_CLASS, Wn.Caption)
    If WinHWnd = 0 Then
    '/////////////////////////////////////////////////////////////////////////////
    '// Windows piles all kinds of messages in the caption in square brackets.  //
    '// [Compatibility Mode], for example...                                    //
    '// Ignore this junk and search for the caption of the passed window.       //
    '// cool.blue@y7mail.com                                                    //
    '/////////////////////////////////////////////////////////////////////////////
        WinHWnd = FindWindowEx(DeskHWnd, 0&, C_EXCEL_WINDOW_CLASS, vbNullString)
        Do While WinHWnd <> 0 And Not isFound
            str = getWindowTextHWnd(WinHWnd)
            If InStr(str, Wn.Caption) Then
                isFound = True
                Exit Do
            Else
                WinHWnd = FindWindowEx(DeskHWnd, WinHWnd, C_EXCEL_WINDOW_CLASS, vbNullString)
            End If
        Loop
    End If
    If Not isFound Then XLHWnd = FindWindowEx(0&, XLHWnd, C_EXCEL_APPLICATION_CLASS, vbNullString)
Loop

If WinHWnd = 0 Then
    ''''''''''''''''''''''''''''''''
    ' an error occurred
    ''''''''''''''''''''''''''''''''
    ErrNum = Err.LastDllError
    ErrText = GetSystemErrorMessageText(ErrorNumber:=ErrNum)
    MsgBox "An error occurred with FindWindow for " & C_EXCEL_WINDOW_CLASS & vbCrLf & _
        "Error: " & CStr(ErrNum) & "  Hex(" & Hex(ErrNum) & ")" & vbCrLf & _
        "Description: " & ErrText
        GetHWndOfWindow = 0
        Exit Function
End If

GetHWndOfWindow = WinHWnd

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
Private Function getWindowTextHWnd(HWnd As LongPtr) As String
Dim txt As String, txtLen As Long

    txtLen = GetWindowTextLength(HWnd) + 1
    txt = String(txtLen, vbNullChar)
    Call getWindowText(HWnd, txt, txtLen)
    getWindowTextHWnd = txt
    
End Function

Private Function className(HWnd As LongPtr) As String
Dim strText As String, strLen As Long

    strText = String$(128, Chr$(0))
    strLen = GetClassName(HWnd, strText, 128)
    className = Left(strText, strLen)
    
End Function
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


