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
Dim Notify As New cActiveSheetBus

'This is needed to boot the active sheet because the
'Worksheet_Activate event does not fire in the sheet
Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    Notify.onWindowActivate ActiveSheet
End Sub
