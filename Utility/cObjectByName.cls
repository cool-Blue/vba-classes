VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cObjectByName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const defDELIM As String = "-"
Private mObject As Object
Private mArgs() As String

Private mProperties As String
Private mDelim As String
Property Let Properties(RHS As String)
  mProperties = RHS
End Property
Private Property Get Args() As String()
  Args = Split(mProperties, mDelim)
End Property

Property Let baseObject(RHS As Object)
  Set mObject = RHS
End Property
Property Let Value(RHS As Variant)
  Call setResolvedObject(mObject, mProperties, RHS)
End Property
Property Get Value() As Variant
  If TypeOf getResolvedValue(mObject, mProperties) Is Object  Then
    Set Value = getResolvedValue(mObject, mProperties)
  Else
    Value = getResolvedValue(mObject, mProperties)
  End If
End Property
Property Get Reference() As String()
'  Set Reference = getResolvedObject(mObject, mProperties)
End Property
Private Function getResolvedValue(ByVal o As Object, _
                                  Properties As String) As Variant
Dim pArgs() As String, p As Variant

  pArgs = Args
  For Each p In pArgs
    If Not setObjectToProperty(o, CStr(p)) Then
      getResolvedValue = CallByName(o, CStr(p), VbGet)
      Exit For
    End If
  Next p
End Function
Private Function getResolvedObject(ByVal o As Object, _
                                  Properties As String) 'As String()
'returns the name of the last two properties as strings
'- incomplete concept would need to create another class to manage this...
'not worth the overhead. Just pass this type of object and use .Value
Dim pArgs() As String, p As Variant
  pArgs = Args
  For Each p In pArgs
    If setObjectToProperty(o, CStr(p)) Then
   '   Set getResolvedObject(0) = CStr(p)
    Else
   '   getResolvedObject(1) = CStr(p)
      Exit For
    End If
  Next p
End Function
Private Sub setResolvedObject(ByVal o As Object, _
                              Properties As String, _
                              Value As Variant)
Dim pArgs() As String, p As Variant, vo As Variant

  pArgs = Args
  Set vo = o
  For Each p In pArgs
    If Not setObjectToProperty(o, CStr(p)) Then
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
Private Function setObjectToProperty(o As Object, p As String) As Boolean
  If TypeOf CallByName(o, p, VbGet) Is Object  Then
    Set o = CallByName(o, p, VbGet)
    setObjectToProperty = True
  Else
    setObjectToProperty = False
  End If
End Function
Private Function sProcedure() As String
  sProcedure = "'" & callBackProcedure & " " & """" & mName & """," & """" & fmtTime(nextTime) & """'"
End Function



Private Sub Class_Initialize()
  mDelim = defDELIM
End Sub
