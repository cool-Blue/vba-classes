VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cldefenSheetObjects_InitDelColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Delegates As Collection
Public Property Get Item(index As Variant) As clSheetObjects_InitDelegate
Attribute Item.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Item.VB_UserMemId = 0
    Set Item = Delegates.Item(index)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Enumerator: Edit in Notepad to see this.
'    Attribute NewEnum.VB_UserMemId = -4
'    Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = Delegates.[_NewEnum]
End Property

Private Sub Class_Initialize()
Dim var As clSheetObjects_InitDelegate, ref As clSheetObjects_InitDelegate
    Set Delegates = New Collection
    With Delegates
        Set var = New clSheetObjects_InitVariant
        .Add var, "Variant"
        Set ref = New clSheetObjects_InitRef
        .Add ref, "Ref"
    End With
End Sub

Private Sub Class_Terminate()
    Set Delegates = Nothing
End Sub
