VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBlockDiagramHeadBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const senderMeForward As String = "Me Forward"
Const senderMeBackward As String = "Me Backward"
Private mDirections As cCollection

Private mLastSender As String
Public mYes As Boolean

Public Property Get Yes() As Boolean
Attribute Yes.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Yes.VB_UserMemId = 0
  Yes = mYes
End Property
Public Property Let Yes(RHS As Boolean)
  mYes = RHS
End Property
Private Sub Class_Initialize()
  Set mDirections = New cCollection
  With mDirections
    .Add False, senderMeForward
    .Add True, senderMeBackward
  End With
  Yes = False
End Sub
Public Property Get sender(Backwards As Boolean) As String
  mLastSender = IIf(Backwards, senderMeBackward, senderMeForward)
  sender = mLastSender
End Property
Property Get LastSender() As String
  LastSender = mLastSender
End Property
Public Property Get Backwards(sender As String) As Boolean
  If mDirections.isIncluded(sender) Then
    Backwards = mDirections(sender)
  Else
    Backwards = False
  End If
End Property
