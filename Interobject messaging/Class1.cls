VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mName As String
Public Function getName() As String
  getName = CallByName(Me, "Name", VbGet)
End Function
Public Function doIt(doWhat As String, firstName As String, _
                     secondName As String) As String
Attribute doIt.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute doIt.VB_UserMemId = 0
Dim params(1) As Variant
  params(0) = firstName
  params(1) = secondName
  doIt = CallByName(Me, doWhat, VbMethod, firstName, secondName)
End Function

Public Property Get Name() As String
  Name = mName
End Property
Public Function setName(firstName As String, _
                         Optional secondName As String)
  mName = firstName & " " & secondName
  setName = mName
End Function

