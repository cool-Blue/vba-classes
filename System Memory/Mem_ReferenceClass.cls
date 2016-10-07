VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mem_ReferenceClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Example Usage:
'    Dim mRef As Mem_ReferenceClass
'    Set mRef = New Mem_ReferenceClass
'    With mRef
'        .initObject 32, obj
'        OR
'        .Object(obj) = 32
'        OR
'        .passPointers Obj, varPtr(Obj), objPtr(Obj), 32
'    End With

Private Enum VarClass
    isValue
    isReference
End Enum
Private mVarStructure As Mem_VariantStructure
Dim mObj As Object
Dim mVar As Variant
Dim mClass As VarClass
Dim mVarPtr As LongPtr
Dim mObjPtr As LongPtr
Dim mBytes As Long
Public varType As String
Public isByRefVariant As Boolean


Public Sub passPointers(obj As Variant, pVar As LongPtr, pObj As LongPtr, byteCount As Long)
    mClass = IIf(IsObject(obj), isReference, isValue)
    mBytes = byteCount
    Select Case mClass
    Case isValue
        mVar = obj
        mVarPtr = pVar
        mObjPtr = mVarPtr
    Case isReference
        Set mObj = obj
        mObjPtr = pObj
    End Select
    
End Sub
Public Sub initObject(ByVal byteCount As Long, ByRef obj As Variant)
    
    mClass = IIf(IsObject(obj), isReference, isValue)
    
    With mVarStructure
        .initVar obj
        DoEvents
        isByRefVariant = ((.vt And &HF000) = &H4000)
        mVarPtr = .IVal
        
        mBytes = byteCount
        Select Case mClass
        Case isValue
            mVar = obj
            mObjPtr = mVarPtr
        Case isReference
            Set mObj = obj
            mObjPtr = .IValpointsTo 'ObjPtr(obj)
            If isByRefVariant Then
                mVarPtr = mObjPtr
                mObjPtr = Mem_Read_Word(mObjPtr)
            Else
                mObjPtr = mObjPtr
            End If
        End Select
    End With 'mVarStructure
    
End Sub
Property Let Object(ByRef obj As Variant, ByVal byteCount As Long)
Attribute Object.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Object.VB_UserMemId = 0

    initObject byteCount, obj
        
End Property
Property Get VarClass() As String
    VarClass = sClass(mClass)
End Property
Property Get Address() As LongPtr
    Address = mVarPtr
End Property
Property Get AddressX() As String
    AddressX = "0x" & HexPtr(mVarPtr)
End Property
Property Get refAddress() As LongPtr
    refAddress = ptrValue(mVarPtr)
End Property
Property Get byteCount() As Long
    byteCount = mBytes
End Property
Property Get ContentsX(Optional byteCount As Variant) As String
    If IsMissing(byteCount) Then byteCount = mBytes
    ContentsX = Mem_ReadHex_Words(mVarPtr, byteCount)
End Property
Property Get pAddress() As LongPtr
    Select Case mClass
    Case isValue
        pAddress = 0
    Case isReference
        pAddress = mObjPtr
    End Select
End Property
Property Get pAddressX() As String
    Select Case mClass
    Case isValue
        pAddressX = ""
    Case isReference
        pAddressX = "0x" & HexPtr(mObjPtr)
    End Select
End Property
Property Get pContentsX(Optional byteCount As Variant) As String
    Select Case mClass
    Case isValue
        pContentsX = ""
    Case isReference
        If IsMissing(byteCount) Then byteCount = mBytes
        pContentsX = Mem_ReadHex_Words(mObjPtr, byteCount)
    End Select
End Property
Private Function sClass(vc As VarClass) As String
    Select Case vc
    Case isValue
        sClass = "Scalar"
    Case isReference
        sClass = "Complex"
    End Select
End Function

Private Sub Class_Initialize()
    Set mVarStructure = New Mem_VariantStructure
End Sub
