VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clSheetObjects_InitRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const mType As String = "Ref"

Implements clSheetObjects_InitDelegate
Private Property Get clSheetObjects_InitDelegate_TypeName() As String
    clSheetObjects_InitDelegate_TypeName = mType
End Property

Private Sub clSheetObjects_InitDelegate_constructVariable(nameRange As Range, typeRange As Range, baseAddressRange As Range, Contents As Range, varRange As Range)

    Set baseAddressRange = nameRange.Offset(1)
    
    Set Contents = baseAddressRange.Offset(0, 1)
    Set varRange = nameRange.Resize(2, 2)

End Sub

