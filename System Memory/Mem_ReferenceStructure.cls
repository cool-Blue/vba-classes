VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mem_ReferenceStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Example Usage:
'    Dim mRef As Mem_ReferenceStructure
'    Set mRef = New Mem_ReferenceStructure
'    With mRef
'        .initObject 32, obj
'        OR
'        .Object(obj) = 32
'        OR
'        .passPointers Obj, varPtr(Obj), objPtr(Obj), 32
'    End With

Public Enum resolveTo
    unResolved
    refResolved
    prefResolved
    varIVal
    varpIVal
End Enum
Public Enum enVarClass
    isValue
    isReference
End Enum

Private mVarStructure As Mem_VariantStructure
Private mObjName As String
Dim mObj As Object
Dim mVar As Variant
Dim mClass As enVarClass
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
Public Sub initObject(ByVal byteCount As Long, ByRef obj As Variant, Optional objName As String = vbNullString)
    
    mObjName = objName
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
'Default member: Edit in Notepad to see this.
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
Property Get ContentsX(Optional byteCount As Variant, Optional delim As String = " ", Optional baseTag As String = "") As String
    If IsMissing(byteCount) Then byteCount = mBytes
    ContentsX = Mem_ReadHex_Words(mVarPtr, byteCount, delim, baseTag)
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
Property Get pContentsX(Optional byteCount As Variant, Optional delim As String = " ", Optional baseTag As String = "") As String
    Select Case mClass
    Case isValue
        pContentsX = ""
    Case isReference
        If IsMissing(byteCount) Then byteCount = mBytes
        pContentsX = Mem_ReadHex_Words(mObjPtr, byteCount, delim, baseTag)
    End Select
End Property
Private Function sClass(vc As enVarClass) As String
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
Public Sub populateRef(wks As Worksheet, spaceName As String, spaceType As String, varName As String, Optional resolveTo As resolveTo = refResolved)
Const delim As String = " "
Dim nsgv As SheetObjects_GlobalRef, sa() As String, byteCount As Long, Addr As String, Cont As String

    Set nsgv = New SheetObjects_GlobalRef
    With nsgv(wks, spaceName, spaceType, varName)
        If .nameRange Is Nothing Then Exit Sub
        
        byteCount = .WordCount * PTR_LENGTH
        Select Case resolveTo
            Case unResolved
                Cont = mVarStructure.baseTagVarX
                Addr = mVarStructure.AddressX
            Case refResolved
                Cont = ContentsX(byteCount:=byteCount, delim:=delim, baseTag:="0x")
                Addr = AddressX
            Case prefResolved
                Cont = pContentsX(byteCount:=byteCount, delim:=delim, baseTag:="0x")
                Addr = pAddressX
            Case varIVal
                Cont = "0x" & mVarStructure.IValpointsToX
                Addr = "0x" & mVarStructure.IValx
            Case varpIVal
                Cont = "0x" & HexPtr(Mem_Read_Word(mVarStructure.IValpointsTo))
                Addr = "0x" & mVarStructure.IValpointsToX
        End Select
        
        .populateRef Addr, Cont
    End With

End Sub
