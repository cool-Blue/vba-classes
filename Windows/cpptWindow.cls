VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cpptWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' cpptWindow
' By Cool Blue cool.blue@y7mail.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Excel window class names
Private Const C_EXCEL_APPLICATION_CLASS = "XLMAIN"
Private Const C_EXCEL_DESKTOP_CLASS = "XLDESK"
Private Const C_EXCEL_WINDOW_CLASS = "EXCEL7"
'Powerpoint window class names
Private Const C_PPT_DESKTOP_CLASS As String = "MDIClient"
Private Const C_PPT_WINDOW_CLASS = "mdiClass"
Private Const C_PPT_SLIDESHOW_CLASS As String = "screenClass"
Private Const C_PPT_PRESENTER As String = "PowerPoint Presenter View"
'VBA window class names
Private Const C_VBA_MAIN As String = "wndclass_desked_gsk"

Private ohWnd As cHWnd
Private PPT As Application
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Handle(Wn)
  Handle = GetHWndOfWindow(Wn)
End Property

Private Property Get pptClassName() As String
  pptClassName = GetPPTClassName
End Property
Private Sub Class_Initialize()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class_Initialize
' This is called when the class is created.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  Set PPT = Application
  Set ohWnd = New cHWnd
  
End Sub

Private Sub Class_Terminate()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class_Terminate
' This is called when the class terminates.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Set ohWnd = Nothing
  Set PPT = Nothing
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private Procedures
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetPPTClassName() As String
  Dim lnVer As Long
  
  lnVer = Val(Application.Version)
  Select Case lnVer
    Case 8:       GetPPTClassName = "PP97FrameClass" 'PPT97
    Case 9 To 12: GetPPTClassName = "PP" & lnVer & "FrameClass"  '9=PPT2K, 10=PPTXP, 11=PPT2003, 12=PPT2007
    Case 14, 15:     GetPPTClassName = "PPTFrameClass"  'PPT2010
    Case Else:    GetPPTClassName = "" '!NewerVersion!
  End Select
  
End Function
Private Function GetHWndOfWindow(Wn)
Const myName As String = "cpptWindow.GetHWndOfWindow"
Dim db As New cDebugReporter

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetHWndOfWindow
' This returns the window handle (HWnd) of the window referenced by Wn.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim WinHWnd As Long

  db.Report caller:=myName, Context:="Class: " & vbTab & pptClassName, Message:="Looking for:" & vbTab & Wn.Caption
  
' Check for a slide show...
'  WinHWnd = oHWnd.FindHWndExFuzzy(W:=Wn, ClassName:=C_PPT_SLIDESHOW_CLASS, _
'                                  NotCaption:=C_PPT_PRESENTER)
' Look for VBA window...
    WinHWnd = ohWnd.FindHWndExFuzzy(W:=Wn, ClassName:=C_VBA_MAIN)

  If WinHWnd = 0 Then
  ' Find the application window
    WinHWnd = ohWnd.FindHWndExFuzzy(ClassName:=pptClassName)
  ' Find the MDI Client
    WinHWnd = ohWnd.FindHWndExFuzzy(HWndParent:=WinHWnd, ClassName:=C_PPT_DESKTOP_CLASS)
  ' Find the active slide
    WinHWnd = ohWnd.FindHWndExFuzzy(HWndParent:=WinHWnd, ClassName:=C_PPT_WINDOW_CLASS)
  End If
  If WinHWnd = 0 Then
      ''''''''''''''''''''''''''''''''
      ' an error occurred
      ''''''''''''''''''''''''''''''''
      'allow caller to check for 0 HWnd
          Exit Function
  End If
  
  db.ExitMessage = "WinHWnd: 0x" & Hex(WinHWnd)
  
  GetHWndOfWindow = WinHWnd

End Function


