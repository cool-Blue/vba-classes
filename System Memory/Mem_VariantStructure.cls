VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mem_VariantStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mVar As Variant
Private mVarX As String
Private ptrToVar As LongPtr
Private mvt As Integer
Private mwReserved1 As Integer
Private mwReserved2 As Integer
Private mwReserved3 As Integer
Private mIVal As LongPtr
Private wUnused As Long
'Decimal Components
Private mScale As Byte
Private mSign As Byte
Private mHi32 As Long
Private mLo32 As Long
Private mMid32 As Long


Property Let var(ByRef v As Variant, ByRef vref As Variant)
Attribute var.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Var.VB_UserMemId = 0'   Usage:
'   Dim vs as New clsVariantType
'   vs.Var(vs) = X
'   where X can be anything, including a constant or even vs
'
'   The vref parameter receives the RHS of the assignment but it is moved to a local variable first so VarPtr(vref) does
'   not return the address of the variable passed in by the caller.
'   The v parameter is in brackets on the LHS of the assignment statement and a proper reference to the
'   variable in the caller is brought in, so VarPtr(v) returns the actual address of the variable passed in by the caller.
    initVar v
End Property
'Variant Pointer Structure
Property Get Address() As LongPtr
    Address = ptrToVar
End Property
Property Get AddressX() As String
    AddressX = HexPtr(Address)
End Property
Property Get VarX() As String
    VarX = mVarX
End Property
'Variant Type
    Property Get vt() As Integer
        vt = mvt
    End Property
    Property Get vtx() As String
        vtx = Right(HexPtr(vt), 4)
    End Property
'Variant Long or LongPtr Components
    Property Get IVal() As LongPtr
        IVal = mIVal
    End Property
    Property Get IValx() As String
        IValx = HexPtr(IVal)
    End Property
    Property Get IValpointsTo() As LongPtr
        IValpointsTo = Mem_Read_Word(mIVal)
    End Property
    Property Get IValpointsToX() As String
        IValpointsToX = HexPtr(IValpointsTo)
    End Property
'Variant Decimal Components
    Property Get Scalex() As String
        Scalex = Right(HexPtr(mScale), 2)
    End Property
    Property Get Signx() As String
        Signx = Right(HexPtr(mSign), 2)
    End Property
    Property Get Hi32x() As String
        Hi32x = HexPtr(mHi32)
    End Property
    Property Get Lo32x() As String
        Lo32x = HexPtr(mLo32)
    End Property
    Property Get Mid32x() As String
        Mid32x = HexPtr(mMid32)
    End Property

Public Sub initVar(ByRef v As Variant)
    If IsObject(v) Then
        Set mVar = v
    Else
        mVar = v
    End If
    ptrToVar = VarPtr(v)
    mVarX = Mem_ReadHex_Words(ptrToVar, 64)
    Mem_Copy mvt, v, 2
    Select Case mvt
        Case vbDecimal
            Mem_Copy mScale, ByVal ptrToVar + 2, 1
            Mem_Copy mSign, ByVal ptrToVar + 3, 1
            Mem_Copy mHi32, ByVal ptrToVar + 4, 4
            Mem_Copy mLo32, ByVal ptrToVar + 8, 4
            Mem_Copy mMid32, ByVal ptrToVar + 12, 4
        Case Else
            Mem_Copy mwReserved1, ByVal ptrToVar + 2, 2
            Mem_Copy mwReserved1, ByVal ptrToVar + 4, 2
            Mem_Copy mwReserved1, ByVal ptrToVar + 6, 2
            Mem_Copy mIVal, ByVal ptrToVar + 8, 4
            Mem_Copy wUnused, ByVal ptrToVar + 12, 4
    End Select
End Sub
