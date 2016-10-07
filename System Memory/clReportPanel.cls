VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clReportPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTxtBox As MSForms.TextBox

Property Get txtBoxLoaded() As Boolean
Attribute txtBoxLoaded.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute txtBoxLoaded.VB_UserMemId = 0
    txtBoxLoaded = Not mTxtBox Is Nothing
End Property
Property Let txtBox(sSheetName As String, tb As String)
Dim ws As Worksheet

    If ws = "" Then
        Set ws = ActiveSheet
    Else
        Set ws = ActiveWorkbook.Sheets(sSheetName)
    End If
    
    Set mTxtBox = ws.OLEObjects(tb).Object
    
End Property
Public Sub reportState(ByRef msg As String, logMode As enLogMode)
    
    With mTxtBox
        Select Case logMode
        Case replaceLog
            .Text = msg
        Case pushLog
            If Len(.Text) = 0 Then .Text = msg Else .Text = msg & rept(vbCrLf, 2) & .Text
        Case addToLog
            .Text = .Text & vbCrLf & msg
        End Select
    End With

End Sub

Private Sub Class_Initialize()
'Default behaviour is to connect to the first TextBox it finds in the Active Sheet
Dim ws As Worksheet, pOLE As OLEObject

    Set ws = ActiveSheet
    For Each pOLE In ws.OLEObjects
        If TypeName(pOLE.Object) = "TextBox" Then
            Set mTxtBox = pOLE.Object
            Exit For
        End If
    Next pOLE
End Sub
