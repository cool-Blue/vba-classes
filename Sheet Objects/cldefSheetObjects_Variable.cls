VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cldefSheetObjects_Variable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mNameRange As Range
Private mVarRange As Range
Private mTypeRange As Range
Private mbaseAddressRange As Range
Private mContents As Range
Private mInitDelegate As cldefenSheetObjects_InitDelColl
Property Let nameRange(nRange As Range)
Dim pWordCount As Long

    Set mNameRange = nRange
    Set mTypeRange = nameRange.Offset(0, 1)
    
    If mNameRange Is Nothing Then Exit Property
    
    mInitDelegate(varType).constructVariable mNameRange, mTypeRange, mbaseAddressRange, mContents, mVarRange

End Property
Property Get nameRange() As Range
    Set nameRange = mNameRange
End Property

Public Property Get Range() As Range
    Set Range = mVarRange
End Property
Public Property Get Name() As String
    If Not (mNameRange Is Nothing) Then Name = mNameRange.Text
End Property

Public Property Get varType() As String
' Form is TypeName X
' Where X is an expression with no spaces which is discarded here
Dim s As String
    If Not (mTypeRange Is Nothing) Then
        s = mTypeRange.Text
        If InStr(s, " ") Then
            varType = Left(s, InStr(s, " ") - 1)
        Else
            varType = s
        End If
    End If
End Property

Public Property Get baseAddress() As String
    If Not (mbaseAddressRange Is Nothing) Then baseAddress = mbaseAddressRange.Text
End Property
Public Property Let baseAddress(RHS As String)
    If Not (mbaseAddressRange Is Nothing) Then mbaseAddressRange.Text = RHS
End Property

Public Property Get Contents() As String()
Attribute Contents.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Contents.VB_UserMemId = 0
Dim i As Long, buff() As String, c As Range
    If WordCount <= 0 Then Exit Property
    ReDim buff(WordCount - 1): i = 0
    For Each c In mContents
        buff(i) = c.Text
        i = i + 1
    Next c
    Contents = buff
End Property
Public Property Let Contents(RHS() As String)
    If Not (mContents Is Nothing) Then mContents.Value2 = RHS
End Property

Public Property Get WordCount() As Long
    If Not (mContents Is Nothing) Then
        WordCount = mContents.Rows.Count
    Else
        WordCount = 0
    End If
End Property

Private Sub Class_Initialize()
    Set mInitDelegate = New cldefenSheetObjects_InitDelColl
End Sub
