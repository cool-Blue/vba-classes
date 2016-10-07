VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clSheetObjects_Variant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mVarRange As Range
Private mNameRange As Range
Private mTypeRange As Range
Private mbaseAddressRange As Range
Private mContents As Range
Private WordCount As Long

'/////////////////////////////////////////////////////////////////////////////////////////////
'/Implimented Proporties
'/////////////////////////////////////////////////////////////////////////////////////////////
Public Sub init(NS As cldefSheetObjects_NameSpace, Name As String)
Dim varType As String

    Set mNameRange = NS.varNameRange(Name)
    Set mbaseAddressRange = mNameRange.Offset(1)
    varType = mNameRange.Offset(0, 1).Text
    If InStr(varType, "Variant") Then
        WordCount = Application.max(Val(Trim(Replace(varType, "Variant", ""))) / 4, 4)
    Else
        WordCount = 1
    End If
    Set mContents = mbaseAddressRange.Offset(0, 1).Resize(WordCount)
    Set mVarRange = mNameRange.Resize(WordCount + 1, 2)

End Sub


Public Property Get Range() As Range
    Set Range = mVarRange
End Property

Public Property Get Name() As String
    Name = mNameRange.Text
End Property

Public Property Get varType() As String
    varType = mTypeRange.Text
End Property

Public Property Get baseAddress() As String
    baseAddress = mbaseAddressRange.Text
End Property
Public Property Let baseAddress(RHS As String)
    mbaseAddressRange.Text = RHS
End Property

Public Property Get Contents() As String()
Dim i As Long, buff() As String
    ReDim buff(WordCount - 1)
    For i = 0 To WordCount - 1
        buff(i) = mContents.Cells(i + 1).Text
    Next i
    Contents = buff
End Property
Public Property Let Contents(RHS() As String)
    mContents.Value2 = RHS
End Property

Public Property Get WordCount() As Long
    WordCount = WordCount
End Property
'/////////////////////////////////////////////////////////////////////////////////////////////
