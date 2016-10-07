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

Public Event notifyActivated(source As String, Backwards As Boolean)
Public Event notifyParent(Parent As String, Activate As Boolean)
Public Event notifyElements(Activate As Boolean, elements As cCollection, sender As String)
Public Event notifyElement(Activate As Boolean, element As String, sender As String)

Private mLastSender As String
Private mLastBackwards As Boolean

Public Sub Activated(source As String, Backwards As Boolean)
  RaiseEvent notifyActivated(source, Backwards)
  mLastBackwards = Backwards
End Sub

Public Sub Parent(p As String, Activate As Boolean)
  RaiseEvent notifyParent(p, Activate)
End Sub

Public Sub elements(Activate As Boolean, bdElements As Variant, sender As String)
Const myName As String = "cNotify.elements"
Dim db As New cDebugReporter
    db.Report Caller:=myName, message:=sender

Dim elements As cCollection
  mLastSender = sender
  Select Case TypeName(bdElements)
  Case "cCollection"
    Set elements = bdElements
    If Not bdElements Is Nothing Then
      elements.dbReport
      RaiseEvent notifyElements(Activate, elements, sender)
    End If
  Case "String"
    db.Report message:=CStr(bdElements)
    RaiseEvent notifyElement(Activate, CStr(bdElements), sender)
  End Select
End Sub
Public Property Get LastSender() As String
  LastSender = mLastSender
End Property
Public Property Get LastBackwards() As Boolean
  LastBackwards = mLastBackwards
End Property
Public Property Let LastBackwards(RHS As Boolean)
  mLastBackwards = RHS
End Property

