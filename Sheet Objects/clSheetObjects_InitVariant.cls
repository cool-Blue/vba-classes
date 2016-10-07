VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clSheetObjects_InitVariant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const mType As String = "Variant"

Implements clSheetObjects_InitDelegate
Private Property Get clSheetObjects_InitDelegate_TypeName() As String
    clSheetObjects_InitDelegate_TypeName = mType
End Property

Private Sub clSheetObjects_InitDelegate_constructVariable(nameRange As Range, typeRange As Range, baseAddressRange As Range, Contents As Range, varRange As Range)

Dim pWordCount As Long

    Set baseAddressRange = nameRange.Offset(1)
    
    If InStr(typeRange.Text, "Variant") Then
        pWordCount = Application.max(Val(Trim(Replace(typeRange.Text, "Variant", ""))) / 4, 4)
    Else
        pWordCount = 1
    End If
    
    Set Contents = baseAddressRange.Offset(0, 1).Resize(pWordCount)
    Set varRange = nameRange.Resize(pWordCount + 1, 2)
    
End Sub

