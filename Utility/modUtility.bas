Attribute VB_Name = "modUtility"
Option Explicit
Const cModuleName As String = "Utility"
Const cModuleIndent As String = 1
'************************************************************************************************************************************************
'cOnTime Support
Public Function Text(Number As Variant, precision As Long) As String
Dim fmt As String
    fmt = "0." & rept("0", precision)
    Text = Application.WorksheetFunction.Text(Number, fmt)
End Function
Public Function rept(ch As String, i As Long) As String
    rept = Application.WorksheetFunction.rept(ch, i)
End Function
Public Function max(v1 As Variant, v2 As Variant) As Variant
    max = IIf(v1 > v2, v1, v2)
End Function
