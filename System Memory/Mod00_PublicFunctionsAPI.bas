Attribute VB_Name = "Mod00_PublicFunctionsAPI"
Option Explicit

Public Enum enEndianNess
    littleEndian = True
    bigEndian = False
End Enum
Const ErrNotDivisible As Long = vbObjectError + 1
Const Err1Message As String = "Mem_ReadHex_Words: Length is not divisible by wordLength"
Const ErrZeroPointer As Long = vbObjectError + 2
Const Err2Message As String = "Mem_ReadHex_Words: wordLength Exceeds LongPtr size."
Const Err3Number As Long = vbObjectError + 3
Const Err3Message As String = "Mem_ReadHex_Words: Illegal Pointer Value: zero pointer not allowed."
Const Err4Number As Long = vbObjectError + 4
Const Err4Message As String = "Mem_ReadHex_Words: ????"

Public Const col1 As Long = 40
Public Const col2 As Long = 53
Public Const col3 As Long = 65
'Utility Objects
Public gf As FormattedString
Public gVarStruct As Mem_VariantStructure
Public gRefStruct As Mem_ReferenceClass

'Mem_Read constantly the Length of 4 Bytes
Public Const coBytes As Long = 4 * 8
'Variables storing ObjPtr and VarPtr during runtime
Public lngPtr_objPtr As LongPtr
Public lngPtr_varPtr As LongPtr
'additional to Class1
Public lngPtr_objcls1 As LongPtr
Public lngPtr_varcls1 As LongPtr
'For ClassA and B
Public lngPtr_objclsA As LongPtr
Public lngPtr_varclsA As LongPtr
Public lngPtr_objclsB As LongPtr
Public lngPtr_varclsB As LongPtr
Public lngPtr_objclsA_B As LongPtr
Public lngPtr_varclsA_B As LongPtr
Public bln_TerminateA As Boolean
Public bln_TerminateB As Boolean
'not in use
'Public lngPtr_ObjPtrMulti1 As Long
'Public lngPtr_ObjPtrMulti2 As Long

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Function printMem_ReferenceClassStructure(obj As Object, objName As String, objVarPtr As LongPtr, objObjPtr As LongPtr, RefStruct As Mem_ReferenceClass, _
                                    f As FormattedString, Optional byteCount As LongPtr = 32) As pointerChainType
Dim mpointerChain As pointerChainType
Dim wordOffset As LongPtr

    With RefStruct
        .initObject byteCount, obj
        printMem_ReferenceClassStructure.plptr = .pAddress
        printMem_ReferenceClassStructure.lptr = .Address
        wordOffset = IIf(.isByRefVariant, 0, 8)
    End With 'RefStruct
    
    With mpointerChain
        .lptr = objVarPtr
        .plptr = Mem_Read_Word(.lptr + wordOffset)
        .pplptr = Mem_Read_Word(.plptr)
        Debug.Print f.Format("VarPtr(" & objName & ")", TypeName(obj), "0x" & HexPtr(.lptr), "0x" & HexPtr(.plptr), _
                                            Mem_Read_WordX(.plptr), Mem_ReadHex_Words(.pplptr, byteCount))
        
        .lptr = objObjPtr
        .plptr = Mem_Read_Word(.lptr)
        .pplptr = Mem_Read_Word(.plptr)
        Debug.Print f.Format("ObjPtr(" & objName & ")", "0x" & HexPtr(.lptr), "", "", "0x" & HexPtr(.plptr), Mem_ReadHex_Words(.plptr, byteCount))
    End With

End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Sub PointerToSomething(strname As String, lng_vPtr As Long, lng_oPtr As Long, Length As Long, Optional obj As Variant)
Dim strvHex As String, stroHex As String, stroTargetHex As String
Dim strvPtr As String, stroPtr As String
Dim strMem_ReadRes_v As String, strMem_ReadRes_o As String

    On Error GoTo Terminate
'   Hex of Ptrs
    strvHex = "0x" & HexPtr(lng_vPtr)
    stroHex = "0x" & HexPtr(lng_oPtr)
    If IsMissing(obj) Then
        stroTargetHex = String(Len("0x" & HexPtr(0)), " ")
    Else
        If IsObject(obj) Then stroTargetHex = "0x" & HexPtr(ObjPtr(obj)) Else stroTargetHex = "0x" & HexPtr(0)
    End If
'   Adjustment objName
    If Len(strname) < col1 Then strname = strname & String(col1 - Len(strname), " ")
'   Mem_Read of VarPtr
    strMem_ReadRes_v = " : 0x" & Mem_ReadHex_Words(lng_vPtr, 4)
    
'   Mem_Read of ObjPtr
    strMem_ReadRes_o = " : " & Mem_ReadHex_Words(lng_oPtr, Length)

    Debug.Print strname & " : " & stroTargetHex & " : " & strvHex & strMem_ReadRes_v & _
                        " : " & stroHex & strMem_ReadRes_o & " :"
    
Exit Sub
Terminate:
    Select Case Err.Number
    Case ErrZeroPointer
        strMem_ReadRes_o = " : Can't be obtained without crash"
        Resume Next
    Case Else
        
    End Select
End Sub
Function STRpointerToSomething(strname As String, lng_vPtr As LongPtr, lng_oPtr As LongPtr, Length As Long) As String
Dim strvHex As String, stroHex As String
Dim strvPtr As String, stroPtr As String
Dim strMem_ReadRes_v As String, strMem_ReadRes_o As String

    On Error GoTo Terminate
'   Hex of Ptrs
    strvHex = HexPtr(lng_vPtr)
    stroHex = HexPtr(lng_oPtr)
'   Adjustment objName
    If Len(strname) < (col3 - col2) Then strname = strname & String((col3 - col2) - Len(strname), " ")
'   Mem_Read of VarPtr
    strMem_ReadRes_v = " : 0x" & Mem_ReadHex_Words(lng_vPtr, 4)
    
'   Mem_Read of ObjPtr
    strMem_ReadRes_o = " : " & Mem_ReadHex_Words(lng_oPtr, Length)

    STRpointerToSomething = strname & " : " & "0x" & strvHex & strMem_ReadRes_v & _
                            " : " & "0x" & stroHex & strMem_ReadRes_o
    
Exit Function
Terminate:
    Select Case Err.Number
    Case ErrZeroPointer
        If Len(strMem_ReadRes_o) = 0 Then
            strMem_ReadRes_o = " : Can't be obtained without crash"
        ElseIf Len(strMem_ReadRes_v) = 0 Then
            strMem_ReadRes_v = " : Can't be obtained without crash"
        Else
            STRpointerToSomething = " : Can't be obtained without crash"
        End If
        Resume Next
    Case Else
        
    End Select
End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Sub Do_HeaderOld(Optional str As String)
Const H1 As String = "Object"
    Debug.Print String(col2 + 3, " ") & H1 & String(col3 - col2 - Len(H1), " ") & ":  VarPtr    :  Contents   :  ObjPtr    : Contents"
    
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Do_Header(Optional str As String)
Const H1 As String = "Object"
    Debug.Print
    Debug.Print gf.Format(str, H1, "VarPtr", "Contents", "ObjPtr", "Contents")
    
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Formatting
Public Sub initUtilityObjects()
Const w1 As Long = 42 '35
Const w2 As Long = 13
Const w3 As Long = 13
Const w4 As Long = 13
Const w5 As Long = 13
Const w6 As Long = 74

    Set gf = New FormattedString
    With gf
        .setColWidths w1, w2, w3, w4, w5, w6 ', colwidth ', colwidth, colwidth, colwidth
        .setIndents 0, 1
    End With

    Set gRefStruct = New Mem_ReferenceClass

End Sub


