VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
Dim blockDiagram As cBlockDiagram

  If gcbd Is Nothing Then Set gcbd = New cCollection
  
  If Not gcbd.isIncluded(Me.Name) Then
    Set blockDiagram = New cBlockDiagram
    blockDiagram.preBuiltWs = Me
    gcbd.Add blockDiagram, Me.Name
  ElseIf gcbd(Me.Name).Count < 2 Then
    gcbd.Remove Me.Name
    Set blockDiagram = New cBlockDiagram
    blockDiagram.preBuiltWs = Me
    gcbd.Add blockDiagram, Me.Name
  End If
End Sub

