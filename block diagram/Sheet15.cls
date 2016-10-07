VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()
Dim cb As MSForms.CommandButton
  With CommandButton1
    If .Caption = "On" Then
      .Caption = "Off"
      Me.designMode = True
    Else
      .Caption = "On"
      Me.designMode = False
    End If
  End With
End Sub

Private Sub Worksheet_Activate()
Dim blockDiagram As cBlockDiagram

  If gcbd Is Nothing Then Set gcbd = New cCollection
  
  If Not gcbd.isIncluded(Me.Name) Then
    Set blockDiagram = New cBlockDiagram
    blockDiagram.ws = Me
    gcbd.Add blockDiagram, Me.Name
  ElseIf gcbd(Me.Name).Count < 2 Then
    gcbd.Remove Me.Name
    Set blockDiagram = New cBlockDiagram
    blockDiagram.preBuiltWs = Me
    gcbd.Add blockDiagram, Me.Name
  End If
End Sub
Property Get isEnabled() As Boolean

  isEnabled = Not gcbd Is Nothing
  If Not isEnabled Then Exit Property
  
  isEnabled = isEnabled And gcbd.isIncluded(Me.Name)
  If Not isEnabled Then Exit Property

  isEnabled = isEnabled And Not gcbd(Me.Name).Count < 2

End Property

Property Let designMode(RHS As Boolean)

  If RHS Then
    gcbd.Remove Me.Name
  Else
    EnableBlocDiagram
  End If
  
End Property
Private Sub EnableBlocDiagram()
Dim blockDiagram As cBlockDiagram

  If gcbd Is Nothing Then Set gcbd = New cCollection
  
  Set blockDiagram = New cBlockDiagram
  blockDiagram.preBuiltWs = Me
  gcbd.Add blockDiagram, Me.Name
  
End Sub

