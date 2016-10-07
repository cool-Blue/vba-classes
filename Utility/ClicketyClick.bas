Attribute VB_Name = "ClicketyClick"
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/  Example usage...
'/  In UserForm Module
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/  Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'/  Const myName As String = "UserForm1" & "." & "UserForm_MouseDown"
'/      Debug.Print myName
'/  End Sub
'/
'/  Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'/  Const myName As String = "UserForm1" & "." & "UserForm_MouseUp"
'/      Debug.Print myName
'/  End Sub
'/  Private Sub UserForm_Click()
'/  Const myName As String = "UserForm1" & "." & "UserForm_Click"
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/      If isDoubleClick Then Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/      Debug.Print myName
'/  End Sub
'/
'/  Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'/  Const myName As String = "UserForm1.UserForm_DblClick"
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/      isClick = False
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/      Debug.Print myName
'/  End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public isClick As Boolean
Declare PtrSafe Function GetDoubleClickTime Lib "user32" () As LongPtr

Public Function isDoubleClick() As Boolean
Dim t As Single
Dim dClickTime As Double

dClickTime = GetDoubleClickTime / 1000
t = Timer
isClick = True
Do
    DoEvents
    isDoubleClick = Not isClick
Loop Until Timer > t + dClickTime Or Timer < t

End Function

Sub showForm()
Dim mUF As UserForm1, wks As Worksheet
    Set mUF = New UserForm1
    mUF.Show
    On Error Resume Next
    Unload mUF
End Sub

