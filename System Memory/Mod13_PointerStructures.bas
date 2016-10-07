Attribute VB_Name = "Mod13_PointerStructures"
Option Explicit

Dim gobj As cldefVariantStructure
Dim gScalar As LongPtr
Dim gwks As Worksheet

Sub testObj()
Dim o As cldefReferrence
Dim vs As cldefVariantStructure
Dim lObj As Worksheet

    Set gwks = Worksheets(1)
    
    Set o = New cldefReferrence
    Set lObj = gwks
    o.initObject 32, lObj
    Set vs = New cldefVariantStructure
    vs.initVar lObj
    
    Debug.Print "Local Pointer: VarPtr: " & HexPtr(varPtr(lObj)) & vbTab & "Contents: " & vbTab & Mem_ReadHex_Words(varPtr(lObj), 4)
    Debug.Print "Local Pointer: ObjPtr: " & HexPtr(objPtr(lObj)) & vbTab & "Contents: " & vbTab & Mem_ReadHex_Words(objPtr(gwks), 32)
    Debug.Print "Global:        VarPtr: " & HexPtr(varPtr(gwks)) & vbTab & "Contents: " & vbTab & Mem_ReadHex_Words(varPtr(gwks), 4)
    Debug.Print "Global:        ObjPtr: " & HexPtr(objPtr(gwks)) & vbTab & "Contents: " & vbTab & Mem_ReadHex_Words(objPtr(gwks), 32)
    Debug.Print
    
    Debug.Print "Local Pointer: VarPtr: " & HexPtr(vs.Address) & vbTab & "Contents: " & vbTab & Mem_ReadHex_Words(vs.Address, 4)
    Debug.Print "Local Pointer: ObjPtr: " & HexPtr(vs.IVal) & vbTab & "Contents: " & vbTab & Mem_ReadHex_Words(vs.IVal, 32)
    Debug.Print
    
    Debug.Print "Var Type: " & o.VarClass & vbTab & o.byteCount
    Debug.Print "Address:  " & o.Address & vbTab & "Contents:  " & o.Contents(32)
    Debug.Print "pAddress: " & o.pAddress & vbTab & "pContents: " & o.pContents(32)

End Sub
Private Sub testObjExample()
Const objBytes As Long = 32
Dim o As cldefReferrence
Dim lObj As Worksheet

    Set gwks = Worksheets(1)
    
    Set lObj = gwks
    Set o = New cldefReferrence
    o.initObject objBytes, lObj
    
    Debug.Print "Address:  " & o.Address & vbTab & "Contents:  " & o.Contents(objBytes)
    Debug.Print "pAddress: " & o.pAddress & vbTab & "pContents: " & o.pContents(objBytes)

End Sub

Sub testVarStructureObject()
Dim lObj As Worksheet
Dim vs As cldefVariantStructure
Dim v As Variant

    Set gwks = Worksheets(1)
    Set lObj = gwks
    Set vs = New cldefVariantStructure
'    vs.Var(v) = 0
    Set v = lObj
    vs.initVar lObj
    Debug.Print
    Debug.Print "Global Address: " & HexPtr(varPtr(gwks)) & vbTab & "Reference: " & HexPtr(objPtr(gwks))
    Debug.Print "Local Ref Address: " & HexPtr(varPtr(lObj)) & vbTab & "Reference: " & HexPtr(objPtr(lObj))
    Debug.Print "Address in Class: " & vs.AddressX & vbTab & "vt: " & vs.vtx & vbTab & "IVal: " & vbTab & vs.IValx
    Debug.Print "Contents: " & vbTab & vs.Varx
End Sub
Sub testVarStructureScalar()
Dim vs As cldefVariantStructure
Dim v As Variant
Dim lScalar As Long

    gScalar = &HABCDABCD
    v = gScalar
    Set vs = New cldefVariantStructure
'    vs.Var(v) = 0
    vs.initVar v
    Debug.Print
    Debug.Print "Global Address: " & HexPtr(varPtr(gScalar))
    Debug.Print "Local Ref Address: " & HexPtr(varPtr(v)) & vbTab & "vs Address: " & HexPtr(varPtr(vs))
    Debug.Print "Address in Class: " & HexPtr(vs.Address) & vbTab & "vt: " & vs.vt & vbTab & "IVal: " & vbTab & HexPtr(vs.IVal)
    Debug.Print "Contents: " & vbTab & vs.Varx
End Sub
Sub testVarStructureDecimal()
Dim vs As cldefVariantStructure
Dim v As Variant

    v = CDec("3.14159265358979323846")  'gScalar
    Set vs = New cldefVariantStructure
'    vs.Var(v) = 0
    vs.initVar v
    Debug.Print
    Debug.Print "Global Address: " & HexPtr(varPtr(gScalar))
    Debug.Print "Local Ref Address: " & HexPtr(varPtr(v)) & vbTab & "vs Address: " & HexPtr(varPtr(vs))
    Debug.Print "Address in Class: " & vs.AddressX & vbTab & "vt: " & vs.vtx & vbTab & "Scale: " & vs.Scalex _
                & vbTab & "Sign: " & vs.Signx & vbTab & "Hi32: " & vs.Hi32x & vbTab & "Lo32: " & vs.Lo32x & vbTab & "Mid32: " & vs.Mid32x
    Debug.Print "Contents: " & vbTab & vs.Varx
End Sub


