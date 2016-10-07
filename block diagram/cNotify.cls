VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event message(MesageType As String, Sender As String, Receiver As Variant, _
                      Action As Variant)

Private mLastSender As String

Public Property Get LastSender() As String
  LastSender = mLastSender
End Property
Public Sub SendMessage(MesageType As String, Sender As String, _
                       Receiver As Variant, Action As Variant)
Attribute SendMessage.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute SendMessage.VB_UserMemId = 0
Const myName As String = "cNotify.SendMessage"
Dim db As New cDebugReporter
    db.Report Caller:=myName

Dim Receivers As cCollection

  mLastSender = Sender
  Select Case TypeName(Receiver)
  Case "cCollection"
    If Not Receiver Is Nothing Then
      Set Receivers = Receiver
      Receivers.dbReport
      RaiseEvent message(MesageType, Sender, Receiver, Action)
    End If
  Case "String"
    db.Report message:=CStr(Receiver)
    RaiseEvent message(MesageType, Sender, Receiver, Action)
  Case Else
    'Ignore
  End Select
End Sub
