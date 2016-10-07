VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBlockDiagramMessageProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Process(MessageType As String, Sender As String, Receiver As Variant, _
                   Action As Variant, Elements As Collection)
Attribute Process.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Process.VB_UserMemId = 0
  CallByName Me, MessageType, VbMethod, Sender, Receiver, Action, Elements

End Sub
Public Sub notifyActivated(source As String, Receiver As Variant, _
                            Backwards As Variant, Elements As Collection)
Const myName As String = "cBlockDiagram.mNotifier_Activated"
Static LastBackwards As Boolean, CurrentBlock As String
Dim db As New cDebugReporter
    db.Report Caller:=myName, Message:="transfering from " & _
                      emptyString(CurrentBlock) & " to " & emptyString(source)
                      

' If source is different from the starting block or the direction is reversed
  If CurrentBlock <> source Or Backwards <> LastBackwards Then
  
'   If there is a current starting block
    If CurrentBlock <> vbNullString Then
'     Erase the current highlights
      Elements(CurrentBlock).headBlock(LastBackwards) = False
    End If
    
    If source <> vbNullString Then
'     Establish a new starting block
      CurrentBlock = source 'Elements(source).Name
      Elements(CurrentBlock).headBlock(CBool(Backwards)) = True
    
    Else
    
      CurrentBlock = vbNullString
    
    End If
    LastBackwards = Backwards
  End If
  
End Sub

Public Sub notifyParent(Sender As String, Parent As Variant, _
                         Activate As Variant, Elements As Collection)
  Elements(Parent).Highlighted = Activate
End Sub

Public Sub notifyElement(Sender As String, element As String, _
                          Activate As Variant, Elements As Collection)
  
  With Elements(element)
    If Activate <> .Activated Then
      .Activated = Activate
      If Activate Then .Activate Sender Else .deActivate Sender
    End If
  End With
  
End Sub

Public Sub notifyElements(Sender As String, bdReceivers As Variant, _
                            Activate As Variant, Elements As Collection)
Const myName As String = "cBlockDiagramMessageProcessor.notifyElements"
Dim db As New cDebugReporter
    db.Report Caller:=myName

Dim Receivers As cCollection

  Select Case TypeName(bdReceivers)
  Case "cCollection"
    If Not bdReceivers Is Nothing Then
      Set Receivers = bdReceivers
      Receivers.dbReport
      Dim s As Variant
      For Each s In Receivers
        notifyElement Sender, CStr(s), Activate, Elements
      Next s
    End If
  Case "String"
    db.Report Message:=CStr(bdReceivers)
    notifyElement Sender, CStr(bdReceivers), Activate, Elements
  End Select

End Sub

Private Function emptyString(s As String) As String
  emptyString = IIf(s = vbNullString, "nothing", s)
End Function


