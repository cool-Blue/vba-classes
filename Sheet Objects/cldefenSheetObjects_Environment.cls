VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cldefenSheetObjects_Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mNameSpaces As Collection
Private nameSpaceNames As Collection
Public Sub Initialize()
    Set mNameSpaces = New Collection
End Sub
Private Sub Class_Initialize()
    Initialize
End Sub

Public Sub Terminate()
    Set mNameSpaces = Nothing
End Sub
Private Sub Class_Terminate()
    Terminate
End Sub

Public Sub Add(ByVal Item As cldefenSheetObjects_NameSpace)
    mNameSpaces.Add Item, Item.Name
End Sub

Public Sub AddSpaces(ParamArray arr() As Variant)
Dim v As Variant, NS As cldefenSheetObjects_NameSpace
    For Each v In arr
        Set NS = v
        mNameSpaces.Add NS, NS.Name
    Next
End Sub

Public Sub Remove(index As Variant)
    mNameSpaces.Remove index
End Sub

Public Property Get Item(index As Variant) As cldefenSheetObjects_NameSpace
Attribute Item.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Item.VB_UserMemId = 0
    Set Item = mNameSpaces.Item(index)
End Property

Public Property Get Count() As Long
    Count = mNameSpaces.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Enumerator: Edit in Notepad to see this.
'    Attribute NewEnum.VB_UserMemId = -4
'    Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mNameSpaces.[_NewEnum]
End Property

