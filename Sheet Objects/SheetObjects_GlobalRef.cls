VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetObjects_GlobalRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const moduleName As String = "SheetObjects_GlobalRef"
Const Err10_NameSpaceNotFound As Long = vbObjectError + 1000 + 10
Const Err10_Message As String = "Variable or Name Space not found"
Private mNameSpace As SheetObjects_NameSpace
Public Function globalVar(wks As Worksheet, spaceName As String, spaceType As String) As SheetObjects_NameSpace
Attribute globalVar.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute globalVar.VB_UserMemId = 0
    Set mNameSpace = New SheetObjects_NameSpace
    If Not mNameSpace.Initialise(wks, spaceType, spaceName) Then GoTo exitFail
    Set globalVar = mNameSpace
    If globalVar Is Nothing Then GoTo exitFail
    Exit Function
exitFail:
    Err.Raise Err10_NameSpaceNotFound, , Err10_Message & vbCrLf & "Name Space: " & spaceName
End Function

