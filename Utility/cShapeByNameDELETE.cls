VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cShapeByNameDELETE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mShape As Shape
Public Properties As String

Property Let Shape(RHS As Shape)
  Set mShape = RHS
End Property
Property Get Shape() As Shape
  Set Shape = mShape
End Property
Property Let ResolvedProperty(RHS As Variant)
  Call setResolvedObject(mShape, Properties, RHS)
End Property
Property Get ResolvedProperty() As Variant
  If TypeOf getResolvedObject(mShape, Properties) Is Object  Then
    Set ResolvedProperty = getResolvedObject(mShape, Properties)
  Else
    ResolvedProperty = getResolvedObject(mShape, Properties)
  End If
End Property
Function getResolvedObject(ByVal o As Object, _
                        Properties As String, _
                        Optional delim As String = "-") As Variant
Dim Args() As String, p As Variant

  Args = Split(Properties, delim)
  For Each p In Args
    If TypeOf CallByName(o, CStr(p), VbGet) Is Object  Then
      Set o = CallByName(o, CStr(p), VbGet)
      Set getResolvedObject = o
    Else
      getResolvedObject = CallByName(o, CStr(p), VbGet)
      Exit For
    End If
  Next p
End Function
Sub setResolvedObject(ByVal o As Object, _
                        Properties As String, _
                        Value As Variant, _
                        Optional delim As String = "-")
Dim Args() As String, p As Variant, vo As Variant

  Args = Split(Properties, delim)
  Set vo = o
  For Each p In Args
    If TypeOf CallByName(o, CStr(p), VbGet) Is Object  Then
      Set o = CallByName(o, CStr(p), VbGet)
    Else
      Exit For
    End If
  Next p
  If TypeOf Value Is Object  Then
    Set o = Value
  Else
    On Error Resume Next
    CallByName o, CStr(p), VbLet, Value
    On Error GoTo 0
  End If
End Sub



