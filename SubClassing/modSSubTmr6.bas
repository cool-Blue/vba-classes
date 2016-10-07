Attribute VB_Name = "modSSubTmr6"
Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modSSubTmr6
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
'
' This project uses the SSubTmr6.dll to do the actual subclassing (VBA is just too slow
' to do the actual subclassing -- it can't cope with  the flood of messages sent by
' Windows). SSubTmr6.dll is a free subclassing component available on the web at
' http://www.vbaccelerator.com/codelib/ssubtmr/ssubtmr.htm.
' Technical notes and documentation are at
' http://www.vbaccelerator.com/home/VB/Code/Libraries/Subclassing/SSubTimer/article.asp
'
' See also http://www.cpearson.com/excel/SubclassingWithSSubtmr6.htm
'
' To install this DLL, do the following, with Excel closed.
'    -  Go to http://www.vbaccelerator.com/codelib/ssubtmr/ssubtmr.htm
'    -   Download the SSubTmr6.zip (for Visual Basic 6) file and store it in the folder of
'       your choice.
'    -  Unzip the file into that folder.
'    -  Copy the file SSubTmr6.DLL to the "C:\Windows\System32" folder.
'    -  Go to the Windows Start Menu, choose Run, enter the following including the quotes
'       and click OK:
'               RegSvr32 "C:\Windows\System32\SSubTmr6.dll"
'    -   Now open this workbook Excel and the VBA Editor. If you see the reference for
'       "vbAccelerator VB6 Subclassing and Timer Assistant":
'               If it is checked, the refrence is intact. Close the References dialog.
'               If it is marked "MISSING", click Browse and navigate to
'                   "C:\Windows\System23\SSubTmr6.dll" and click OK.
'       If you don't see the reference for "vbAccelerator VB6 Subclassing and Timer Assistant",
'       scroll down in the list until you find it and put a check next to it. If you don't
'       find it in the list, click Browse and navigate to "C:\Windows\System23\SSubTmr6.dll"
'       and click OK.
'
' Once the refrence is in place, you can use the SSubTmr6 component.
' In the object modules (a class module, ThisWorkbook, or a UserForm code module) in which you
' want to do the subclassing, first Implement the ISubClass interface:
'
'           Implements SSubTimer6.ISubclass
'
' Then add the three members of this interface:
'
'       Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
'           ' This Property statement is not used by must be included in the code. This is
'           ' a requirement of using Implements.
'       End Property
'
'       Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
'           ' This property should return the emrPreprocess constant.
'           ISubclass_MsgResponse = SSubTimer6.emrPreprocess
'       End Property
'
'       Private Function ISubclass_WindowProc(ByVal Hwnd As Long, ByVal iMsg As Long, _
'                ByVal WParam As Long, ByVal LParam As Long) As Long
'            ' This procedure is called automatically when there is a message for the HWnd
'            ' specified in the AttachMessage function. It is in this procedure that you
'            ' test HWnd, iMsg, WParam, and LParam. These indicate the window being
'            ' subclassed, the windows message number, and the WParam and LParam values.
'       End Function
'
' To subclass a particular HWnd for a particular message, you need to call the AttachMessage. E.g.,
'
'        Dim WinHWnd As Long
'        WinHWnd = GetHWndOfWindow(Application.ActiveWindow)
'        SSubTimer6.AttachMessage iwp:=Me, Hwnd:=WinHWnd, iMsg:=WM_VSCROLL
'
'        You can call SSubTimer6.AttachMessage for as many HWnd/iMsg pairs as you
'        like. Be sure to store these values in variables because you'll need them to terminate
'        subclassing. For every call to AttachMessage, there must be a corresponding
'        call to DetachMessage.
'
' This code starts subclassing the window WinHWnd for the message WM_VSCROLL.
' Once you have started subclassing the window, the ISubclass_WindowProc procedure will
' automatically be called when there is a message to process.
'
' To stop subclassing an HWnd/Message pair, call the DetachMessage function:
'       SSubTimer6.DetachMessage iwp:=Me, Hwnd:=WinHWnd, iMsg:=WM_VSCROLL
' There must be one call to DetachMessage for each call to AttachMessage.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

