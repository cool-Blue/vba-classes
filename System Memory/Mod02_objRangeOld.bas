Attribute VB_Name = "Mod02_objRange"
Option Explicit
'Reference to the object Range which already exists
    'objRange is assigned a different pointer
    'The macro shows a difference in Mem_ReadHex - Assuming by the different reading
    'objRange should be cleared from memory

Private Sub Test_rng()
Const myName As String = "Test_rng"
Const myType As String = "Sub"
Dim objRange As Range, mtempLptr As LongPtr, mpointerChain As pointerChainType
Dim objName As String, wordOffset As LongPtr
    
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    objName = "Sheet1.Range(""A1"")"
    
    Call printMem_ReferenceStructure(Sheet1.Range("A1"), "Sheet1.RangeA1", VarPtr(Sheet1.Range("A1")), ObjPtr(Sheet1.Range("A1")), gRefStruct, gf)
    With gRefStruct
        .initObject 32, Sheet1.Range("A1")
        wordOffset = IIf(.isByRefVariant, 8, 0)
    End With 'gRefStruct
    
    Debug.Print gf.Format("VarPtr(" & objName, typeName(Sheet1.Range("A1")), "0x" & HexPtr(VarPtr(Sheet1.Range("A1"))), _
                            "0x" & HexPtr(Mem_Read_Word(VarPtr(Sheet1.Range("A1")) + wordOffset)), _
                            "0x" & HexPtr(Mem_Read_Word(Mem_Read_Word(VarPtr(Sheet1.Range("A1")) + wordOffset))), _
                            Mem_ReadHex_Words(Mem_Read_Word(Mem_Read_Word(VarPtr(Sheet1.Range("A1")) + wordOffset)), 32))
    
    Debug.Print gf.Format("ObjPtr(" & objName & ")", "0x" & HexPtr(ObjPtr(Sheet1.Range("A1"))), "", _
                            "", "0x" & HexPtr(Mem_Read_Word(ObjPtr(Sheet1.Range("A1")))), _
                            Mem_ReadHex_Words(Mem_Read_Word(ObjPtr(Sheet1.Range("A1"))), 32))
        
    PointerToSomething objName, VarPtr(Sheet1.Range("A1")), ObjPtr(Sheet1.Range("A1")), coBytes, Sheet1.Range("A1")
    
    With gRefStruct
        .populateRef ActiveSheet, myName, myType, objName
        Debug.Print gf.Format(objName & "*", "", .AddressX, .pAddressX, .pAddressX, .pContentsX)
    End With 'gRefStruct

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    objName = "Sheets(1).Range(""A1"")"
    
    Call printMem_ReferenceStructure(Sheets(1).Range("A1"), "Sheets(1).RangeA1", VarPtr(Sheets(1).Range("A1")), ObjPtr(Sheets(1).Range("A1")), gRefStruct, gf)
    With gRefStruct
        .initObject 32, Sheets(1).Range("A1")
        wordOffset = IIf(.isByRefVariant, 8, 0)
    End With 'gRefStruct
    
    Debug.Print gf.Format("VarPtr(" & objName, typeName(Sheets(1).Range("A1")), "0x" & HexPtr(VarPtr(Sheets(1).Range("A1"))), _
                            "0x" & HexPtr(Mem_Read_Word(VarPtr(Sheets(1).Range("A1")) + wordOffset)), _
                            "0x" & HexPtr(Mem_Read_Word(Mem_Read_Word(VarPtr(Sheets(1).Range("A1")) + wordOffset))), _
                            Mem_ReadHex_Words(Mem_Read_Word(Mem_Read_Word(VarPtr(Sheets(1).Range("A1")) + wordOffset)), 32))
    
    Debug.Print gf.Format("ObjPtr(" & objName & ")", "0x" & HexPtr(ObjPtr(Sheets(1).Range("A1"))), "", _
                            "", "0x" & HexPtr(Mem_Read_Word(ObjPtr(Sheets(1).Range("A1")))), _
                            Mem_ReadHex_Words(Mem_Read_Word(ObjPtr(Sheets(1).Range("A1"))), 32))
        
    PointerToSomething objName, VarPtr(Sheets(1).Range("A1")), ObjPtr(Sheets(1).Range("A1")), coBytes, Sheets(1).Range("A1")
    
    With gRefStruct
        .populateRef ActiveSheet, myName, myType, objName
        Debug.Print gf.Format(objName & "*", "", .AddressX, .pAddressX, .pAddressX, .pContentsX)
    End With 'gRefStruct

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Set objRange = Sheet1.Range("A1")
objName = "objRange"

lngPtr_objPtr = ObjPtr(objRange)
lngPtr_varPtr = VarPtr(objRange)
    
    Call printMem_ReferenceStructure(objRange, "objRange", VarPtr(objRange), ObjPtr(objRange), gRefStruct, gf)
    
    With gRefStruct
        .initObject 32, objRange
'        lngPtr_objPtr = .pAddress
'        lngPtr_varPtr = .Address
        wordOffset = IIf(.isByRefVariant, 0, 8)
    End With 'gRefStruct
    
    With mpointerChain
        .lptr = VarPtr(objRange)
        .plptr = Mem_Read_Word(.lptr + wordOffset)
        .pplptr = Mem_Read_Word(.plptr)
        Debug.Print gf.Format("VarPtr(objRange)", typeName(objRange), "0x" & HexPtr(.lptr), _
                                "0x" & HexPtr(.plptr), "0x" & HexPtr(.pplptr), Mem_ReadHex_Words(.pplptr, 32))
        
        .lptr = ObjPtr(objRange)
        .plptr = Mem_Read_Word(.lptr)
        .pplptr = Mem_Read_Word(.plptr)
        Debug.Print gf.Format("ObjPtr(objRange)", "0x" & HexPtr(.lptr), "", _
                                "", "0x" & HexPtr(.plptr), Mem_ReadHex_Words(.plptr, 32))
    End With
    
    PointerToSomething "objRange", lngPtr_varPtr, lngPtr_objPtr, coBytes, objRange
    
    With gRefStruct
        .populateRef ActiveSheet, myName, myType, objName
'        .populateRef ActiveSheet, myName, myType, "p" & objName, True
        Debug.Print gf.Format("objRange*", "", .AddressX, .pAddressX, .pAddressX, .pContentsX)
    End With 'gRefStruct

'Set objRange = Nothing

If objRange Is Nothing Then
    Debug.Print vbTab & "Object is nothing"
Else
    Debug.Print vbTab & "Object wasn't cleared before leaving Sub"
End If

    PointerToSomething "objRange", lngPtr_varPtr, ObjPtr(objRange), coBytes
    
End Sub

Sub Run02_rngTest()

    initUtilityObjects
    Do_Header "Running Test_Range..."
    Test_rng
        DoEvents
        Debug.Print "After Runtime"
        PointerToSomething "RangeA1", VarPtr(Sheets(1).Range("A1")), ObjPtr(Sheets(1).Range("A1")), coBytes, Sheets(1).Range("A1")
        PointerToSomething "objRange", lngPtr_varPtr, lngPtr_objPtr, coBytes
    
    With gRefStruct
        .initObject 8 * 4, Sheet1.Range("A1")
        Debug.Print gf.Format("RangeA1*", "", .AddressX, .pAddressX, .pAddressX, .pContentsX)
        .populateRef ActiveSheet, "Global", "Module", "no name"
    End With 'gRefStruct
    
    Debug.Print gf.Format("XxxPtr(Sheets(1).Range(""A1""))", "", _
                            "0x" & HexPtr(VarPtr(Sheets(1).Range("A1"))), Mem_Read_WordX(VarPtr(Sheets(1).Range("A1")) + 8), _
                            "0x" & HexPtr(ObjPtr(Sheets(1).Range("A1"))), Mem_ReadHex_Words(Mem_Read_Word(ObjPtr(Sheets(1).Range("A1"))), 32))

End Sub

Sub HexToDec()
Dim x As Long
x = Val("&H09000000")
End Sub
