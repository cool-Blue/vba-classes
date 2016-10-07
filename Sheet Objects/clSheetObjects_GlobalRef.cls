VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clSheetObjects_GlobalRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mNameSpace As cldefenSheetObjects_NameSpace
Public Function globalVar(wks As Worksheet, spaceName As String, spaceType As String, varName As String) As cldefSheetObjects_Variable
    Set mNameSpace = New cldefenSheetObjects_NameSpace
    mNameSpace.nameSpace wks, spaceType, spaceName
    Set globalVar = mNameSpace.Variable(varName)
End Function

