Attribute VB_Name = "Mod01_objwks"
Option Explicit

Private Sub Test_wks()
'Reference to the object Worksheet which already exists
    'objwks is assigned the same pointer as the object itself
    'The macro can't provide particular information to whether or not
    'objwks got destroyed
' But, objwks is a only scalar pointer
Dim objwks As Worksheet

    PointerToSomething "Sheet1", VarPtr(Sheet1), ObjPtr(Sheet1), coBytes, Sheet1
    
    With gRefStruct
        .initObject 32, Sheet1
        Debug.Print gf.Format("Sheet1*", .pAddressX, .AddressX, .pAddressX, .pAddressX, .pContentsX)
    End With 'gRefStruct
    
Set objwks = Sheet1
lngPtr_objPtr = ObjPtr(objwks)
lngPtr_varPtr = VarPtr(objwks)
    
    PointerToSomething "objwks", lngPtr_varPtr, lngPtr_objPtr, coBytes, objwks
    
    With gRefStruct
        .initObject 32, objwks
        Debug.Print gf.Format("objwks*", .pAddressX, .AddressX, .ContentsX, .pAddressX, .pContentsX)
    End With 'gRefStruct
    
Set objwks = Nothing

If objwks Is Nothing Then
    Debug.Print vbTab & "Object is nothing"
Else
    Debug.Print vbTab & "Object wasn't cleared before leaving Sub"
End If

    PointerToSomething "objwks", lngPtr_varPtr, ObjPtr(objwks), coBytes, objwks

End Sub
Sub Run01_wksTest()

    initUtilityObjects
    Do_Header "Running Test_Worksheet..."
    Test_wks
        DoEvents
        Do_Header "After Runtime..."
        PointerToSomething "Sheet1", VarPtr(Sheet1), ObjPtr(Sheet1), coBytes, Sheet1
        PointerToSomething "objwks", lngPtr_varPtr, lngPtr_objPtr, coBytes

End Sub

