VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Set gheaderPics = Nothing
    Set gfloatHeaders = Nothing
End Sub

Private Sub Workbook_Open()
    Set gheaderPics = New floatingHeaders
    Set gfloatHeaders = New cfloatHeaders
End Sub

Sub testopen()
    Set gheaderPics = New floatingHeaders
    Set gfloatHeaders = New cfloatHeaders
End Sub
Sub testclose()
    Set gheaderPics = Nothing
    Set gfloatHeaders = Nothing
End Sub

