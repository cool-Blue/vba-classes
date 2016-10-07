VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cldefReferrence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Enum VarClass
    isScalar
    isComplex
End Enum
Private mVarStructure As clVariantStructure
Dim mObj As Object
Dim mVar As Variant
Dim mClass As VarClass
Dim mVarPtr As LongPtr
Dim mObjPtr As LongPtr
Dim mBytes As Long
Public varType As String

Public Sub initObject(ByVal byteCount As Long, ByRef obj As Variant)
    
    mClass = IIf(IsObject(obj), isComplex, isScalar)
    
    With mVarStructure
        .initVar obj
        mVarPtr = .IVal
    End With 'mVarStructure
    
    mBytes = byteCount
    Select Case mClass
    Case isScalar
        mVar = obj
        mObjPtr = mVarPtr
    Case isComplex
        Set mObj = obj
        mObjPtr = ObjPtr(obj)
    End Select
    
End Sub
Property Get VarClass() As String
    VarClass = sClass(mClass)
End Property
Property Get Address() As String
    Address = "0x" & HexPtr(mVarPtr)
End Property
Property Get byteCount() As Long
    byteCount = mBytes
End Property
Property Get Contents(Optional byteCount As Variant) As String
    If IsMissing(byteCount) Then byteCount = mBytes
    Contents = Mem_ReadHex_Words(mVarPtr, byteCount)
End Property
Property Get pAddress() As String
    Select Case mClass
    Case isScalar
        pAddress = ""
    Case isComplex
        pAddress = "0x" & HexPtr(mObjPtr)
    End Select
End Property
Property Get pContents(Optional byteCount As Variant) As String
    Select Case mClass
    Case isScalar
        pContents = ""
    Case isComplex
        If IsMissing(byteCount) Then byteCount = mBytes
        pContents = Mem_ReadHex_Words(mObjPtr, byteCount)
    End Select
End Property
Private Function sClass(vc As VarClass) As String
    Select Case vc
    Case isScalar
        sClass = "Scalar"
    Case isComplex
        sClass = "Complex"
    End Select
End Function

Private Sub Class_Initialize()
    Set mVarStructure = New clVariantStructure
End Sub
