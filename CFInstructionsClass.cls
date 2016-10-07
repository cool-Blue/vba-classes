VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFInstructionsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const IndxCol = 1
Const ActionCol = 8
Const p1Col = 9
Dim Instructions As Range
Public InstCount As Long

Dim target As Range, ws As Worksheet
Dim fcs As FormatConditions, fc As FormatCondition
Public instPointer As Long

Property Get Indx(i As Long) As Long
    Indx = Instructions.Cells(i, IndxCol)
End Property
Property Let Indx(i As Long, newIndx As Long)
    Instructions.Cells(i, IndxCol) = newIndx
End Property

Property Get Action(i) As String
    Action = Instructions.Cells(i, ActionCol)
End Property
Property Get p1(i As Long) As Variant
    p1 = Instructions.Cells(i, p1Col)
End Property
Private Sub Class_Initialize()
    Set Instructions = Application.InputBox("Select Instruction range", "Manage Conditional Formatting", , , , , , 8)
    Set target = Application.InputBox("Select target range", "Manage Conditional Formatting", , , , , , 8)
    Set fcs = target.FormatConditions
    Set ws = target.Worksheet
    InstCount = Instructions.Rows.Count
    instPointer = 1
End Sub
Public Sub Delete(j As Long)
Dim i As Long

    fcs(j).Delete
    For i = 1 To InstCount
        If Indx(i) > j Then
            Indx(i) = Indx(i) - 1
        End If
    Next i
End Sub
Public Function doAction() As Boolean
Dim r As Range, i As Long, h As Long

    If instPointer < 1 Or instPointer > InstCount Then GoTo exitFalse
    i = instPointer
    doAction = True
    Select Case Action(i)
    Case "DELETE"
        On Error GoTo exitFalse
        Delete Indx(i)
        On Error GoTo 0
    Case "NEWENDROW"
'        Set fc = fcs(i)
        Set r = fcs(Indx(i)).AppliesTo
        h = p1(i) - r.Row + 1
        If h > 0 Then
            Set r = r.Resize(h)
            fcs(Indx(i)).ModifyAppliesToRange r
        End If
    Case "NEWSTARTROW"
'        Set fc = fcs(i)
        Set r = fcs(Indx(i)).AppliesTo
        h = r.Rows.Count + r.Row - p1(i)
        If h > 0 Then
            Set r = r.Offset(p1(i) - r.Row, 0).Resize(h, r.Columns.Count)
            fcs(Indx(i)).ModifyAppliesToRange r
        End If
    Case Else 'skip unkown actions
    End Select
    instPointer = instPointer + 1
    Exit Function
exitFalse:
    doAction = False
End Function
