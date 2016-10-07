VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cannEaseOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const PI As Double = 3.14159265358979
Public startValue As Integer
Public EndValue As Integer
Public dt As Long
Public T As Integer
Private stepCount As Integer

Public Function dCosT() As Integer
Dim dy As Integer
  dy = Round((EndValue - startValue) * PI / (2 * T) * Sin(PI / (2 * T) * stepCount * dt) * dt, 0)
  stepCount = stepCount + 1
'  Debug.Print stepCount, dy
  dCosT = dy
End Function

Private Sub Class_Initialize()
  stepCount = 0
End Sub
