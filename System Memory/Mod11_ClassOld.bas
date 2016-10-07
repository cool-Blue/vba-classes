Attribute VB_Name = "Mod11_Class"
Option Explicit

'Creates a new instance of Class1
    'References to the objects can't be destroyed - even not manually -
    'because the garbage collector believes that the objects are still in use
Private Sub Test_Class()
Const myName As String = "Test_Class"
Const myType As String = "Sub"
Const spaceIndent As Long = 0
Dim objcls1 As Class1, objName As String
Dim pV As LongPtr, pVContents As String, pO As LongPtr, pOContents As String, pIVal As LongPtr, ppIVAL As LongPtr, pppIVal As LongPtr
Dim nsgv As SheetObjects_GlobalRef, dataCrncy As enDataCurrency

Set objcls1 = New Class1
With objcls1
    .Name = "objcls1"
End With

    objName = "objcls1"
    pVContents = Mem_ReadHex_Words(VarPtr(objcls1), coBytes)
    pOContents = Mem_ReadHex_Words(ObjPtr(objcls1), coBytes)
    pV = VarPtr(objcls1): pO = ObjPtr(objcls1)
    lngPtr_objPtr = pO
    lngPtr_varPtr = pV
    
    PointerToSomething objName, pV, pO, coBytes, objcls1, Left(pVContents, 8), pOContents

    Set nsgv = New SheetObjects_GlobalRef
    With nsgv(ActiveSheet, myName, myType)
        .Variable("VarPtr_" & objName).Output HexPtr(pV), pVContents
        .Variable("ObjPtr_" & objName).Output HexPtr(pO), pOContents
    End With
    
    dataCrncy = dcSnapShot
    With gRefStruct
        .initObject 32, objcls1
        .populateSheetObject ActiveSheet, myName, myType, objName, refResolved
        .populateSheetObject ActiveSheet, myName, myType, "p" & objName, prefResolved
        .populateSheetObject ActiveSheet, myName, myType, "raw_" & objName, unResolved, dataCrncy
        .populateSheetObject ActiveSheet, myName, myType, "IVal_" & objName, varIVal
        .populateSheetObject ActiveSheet, myName, myType, "Vt_" & objName, varVt, dataCrncy
        Debug.Print gf.Format(objName & "*", "", .AddressX, .pAddressX, .pAddressX, .pContentsX(coBytes))
    End With 'gRefStruct

'Set objcls1 = Nothing

If objcls1 Is Nothing Then
    Debug.Print "Object is set to nothing"
Else
    Debug.Print "Object wasn't cleared before leaving Sub"
End If

    Debug.Print timeStamp(newLine:=After, Indent:=spaceIndent, caller:=myName, context:="Globals", contextCol:=col1, _
                            message:=STRpointerToSomething("objcls1", lngPtr_varPtr, ObjPtr(objcls1), coBytes), messageCol:=col2)
    
End Sub

Sub Run11_clsTest()
Const myName As String = "Run11_clsTest"
Const myType As String = "Sub"
Const spaceIndent As Long = 0
Dim objName As String
Dim nsgv As SheetObjects_GlobalRef, dataCrncy As enDataCurrency
Dim objcls2 As Class1
Dim objcls3 As Class1
Dim pV As LongPtr, pVContents As String, pO As LongPtr, pOContents As String

    initUtilityObjects

    Debug.Print timeStamp(newLine:=Before, caller:=myName, context:="Start")
    
    Test_Class
    DoEvents
    
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Re-check pointers for objRange
    dataCrncy = dcFresh
    
    objName = "objcls1"
    With gRefStruct
        .populateSheetObject ActiveSheet, myName, myType, objName, refResolved
        .populateSheetObject ActiveSheet, myName, myType, "p" & objName, prefResolved
        .populateSheetObject ActiveSheet, myName, myType, "raw_" & objName, unResolved, dataCrncy
        .populateSheetObject ActiveSheet, myName, myType, "IVal_" & objName, varIVal, dataCrncy
        Debug.Print gf.Format("objRange*", "", .AddressX, .pAddressX, .pAddressX, .pContentsX(coBytes))
    End With 'gRefStruct

    Set nsgv = New SheetObjects_GlobalRef
    nsgv(ActiveSheet, myName, myType)("VarPtr_" & objName).Output HexPtr(lngPtr_varPtr), Mem_ReadHex_Words(lngPtr_varPtr, coBytes)
    nsgv(ActiveSheet, myName, myType)("ObjPtr_" & objName).Output HexPtr(lngPtr_objPtr), Mem_ReadHex_Words(lngPtr_objPtr, coBytes)

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/ Create a local class object
    Set objcls3 = New Class1
    objName = "objcls3"
    objcls3.Name = objName
        
    pVContents = Mem_ReadHex_Words(VarPtr(objcls3), coBytes)
    pOContents = Mem_ReadHex_Words(ObjPtr(objcls3), coBytes)
    pV = VarPtr(objcls3): pO = ObjPtr(objcls3)
    
    PointerToSomething objName, pV, pO, coBytes, objcls3, Left(pVContents, 8), pOContents

    Set nsgv = New SheetObjects_GlobalRef
    With nsgv(ActiveSheet, myName, myType)
        .Variable("VarPtr_" & objName).Output HexPtr(pV), pVContents
        .Variable("ObjPtr_" & objName).Output HexPtr(pO), pOContents
    End With
    
    dataCrncy = dcSnapShot
    With gRefStruct
        .initObject 32, objcls3
        .populateSheetObject ActiveSheet, myName, myType, "p" & objName, prefResolved
        Debug.Print gf.Format(objName & "*", "", .AddressX, .pAddressX, .pAddressX, .pContentsX(coBytes))
    End With 'gRefStruct
    Debug.Print "objcls3.Name: " & objcls3.Name
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/ Investigate the structure of a custom class here...
    Set objcls3 = New Class1
    objcls3.Name = "objcls3"
    Set objcls2 = New Class1
    objName = "objcls2"
    objcls2.Name = objName
        
    pVContents = Mem_ReadHex_Words(VarPtr(objcls2), coBytes)
    pOContents = Mem_ReadHex_Words(ObjPtr(objcls2), coBytes)
    pV = VarPtr(objcls2): pO = ObjPtr(objcls2)
    
    PointerToSomething objName, pV, pO, coBytes, objcls2, Left(pVContents, 8), pOContents

    Set nsgv = New SheetObjects_GlobalRef
    With nsgv(ActiveSheet, myName, myType)
        .Variable("VarPtr_" & objName).Output HexPtr(pV), pVContents
        .Variable("ObjPtr_" & objName).Output HexPtr(pO), pOContents
    End With
    
    dataCrncy = dcSnapShot
    With gRefStruct
        .initObject 32, objcls2
        .populateSheetObject ActiveSheet, myName, myType, "p" & objName, prefResolved
        Debug.Print gf.Format(objName & "*", "", .AddressX, .pAddressX, .pAddressX, .pContentsX(coBytes))
    End With 'gRefStruct
    Debug.Print "objcls2.Name: " & objcls2.Name
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/ Try to point objcls2 to the old objcls1 structure...
    
    Mem_Copy objcls2, lngPtr_objPtr, PTR_LENGTH
    
    objName = "objcls2->objcls1"
    
    pVContents = Mem_ReadHex_Words(VarPtr(objcls2), coBytes)
    pOContents = Mem_ReadHex_Words(ObjPtr(objcls2), coBytes)
    pV = VarPtr(objcls2): pO = ObjPtr(objcls2)
    
    PointerToSomething objName, pV, pO, coBytes, objcls2, Left(pVContents, 8), pOContents

    Set nsgv = New SheetObjects_GlobalRef
    With nsgv(ActiveSheet, myName, myType)
        .Variable("VarPtr_" & objName).Output HexPtr(pV), pVContents
        .Variable("ObjPtr_" & objName).Output HexPtr(pO), pOContents
    End With
    
    dataCrncy = dcSnapShot
    With gRefStruct
        .initObject 32, objcls2
        .populateSheetObject ActiveSheet, myName, myType, "p" & objName, prefResolved
        Debug.Print gf.Format(objName & "*", "", .AddressX, .pAddressX, .pAddressX, .pContentsX(coBytes))
    End With 'gRefStruct
    Debug.Print "objcls2.Name: " & objcls2.Name
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Debug.Print "After Runtime"
    Debug.Print timeStamp(newLine:=No, Indent:=spaceIndent, caller:=myName, context:="Globals", contextCol:=col1, _
                            message:=STRpointerToSomething("objcls1", lngPtr_varPtr, lngPtr_objPtr, coBytes), messageCol:=col2)
    Debug.Print timeStamp(newLine:=No, Indent:=spaceIndent, caller:=myName, context:="Class1", contextCol:=col1, _
                            message:=STRpointerToSomething("objcls1", lngPtr_varcls1, lngPtr_objcls1, coBytes), messageCol:=col2)
    'pointerToSomething "objclsA_B", lngPtr_varclsA_B, lngPtr_objclsA_B, coBytes
    'pointerToSomething "objclsAB", lngPtr_varPtr, lngPtr_objPtr, coBytes
End Sub

