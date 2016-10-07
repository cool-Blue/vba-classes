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
Implements clSheetObjects_Variable

Private mVarRange As Range
Private mNameRange As Range
Private mTypeRange As Range
Private mbaseAddressRange As Range
Private mContents As Range
Private mNameSpace As cldefSheetObjects_NameSpace
Private WordCount As Long
'/////////////////////////////////////////////////////////////////////////////////////////////
'/Implimented Proporties
'/////////////////////////////////////////////////////////////////////////////////////////////
Private Sub clSheetObjects_Variable_init(NS As cldefSheetObjects_NameSpace, Name As String)
Dim varType As String

    Set mNameSpace = NS
    Set mNameRange = mNameSpace.varNameRange(Name)
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

Private Property Get clSheetObjects_Variable_NameSpaceObject() As cldefSheetObjects_NameSpace
    Set clSheetObjects_Variable_NameSpaceObject = mNameSpace
End Property

Private Property Get clSheetObjects_Variable_Range() As Range
    Set clSheetObjects_Variable_Range = mVarRange
End Property

Private Property Get clSheetObjects_Variable_Name() As String
    clSheetObjects_Variable_Name = mNameRange.Text
End Property

Private Property Get clSheetObjects_Variable_varType() As String
    clSheetObjects_Variable_varType = mTypeRange.Text
End Property

Private Property Get clSheetObjects_Variable_baseAddress() As String
    clSheetObjects_Variable_baseAddress = mbaseAddressRange.Text
End Property
Private Property Let clSheetObjects_Variable_baseAddress(RHS As String)
    mbaseAddressRange.Text = RHS
End Property

Private Property Get clSheetObjects_Variable_Contents() As String()
Dim i As Long, buff() As String
    ReDim buff(WordCount - 1)
    For i = 0 To WordCount - 1
        buff(i) = mContents.Cells(i + 1).Text
    Next i
    clSheetObjects_Variable_Contents = buff
End Property
Private Property Let clSheetObjects_Variable_Contents(RHS() As String)
    mContents.Value2 = RHS
End Property

Private Property Get clSheetObjects_Variable_WordCount() As Long
    clSheetObjects_Variable_WordCount = WordCount
End Property
'/////////////////////////////////////////////////////////////////////////////////////////////
