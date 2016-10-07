VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNotifyTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Message(MesageType As String, Sender As String, Receiver As Variant, _
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

  mLastSender = Sender
  
End Sub
