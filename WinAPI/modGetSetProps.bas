Attribute VB_Name = "modGetSetProps"
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modGetSetProps
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' This module contains functions for adding and retrieving property values (Long
' data types) of a window, typically the Excel application's main window. These
' values will remain accessible even when the workbook that created them is closed.
' They will be accessible from any code in any workbook as long as the window exists.
' Usually, you will want to use the Excel main application window (the default for all
' procedures) to store the properties. These properties will persist until Excel closes.
'
' Note that the property can contain only Long data values.
'
' This module contains the following Public procedures (not including Private
' support procedures):
'
'       GetAllProperties - This populates an array of CPropType classes,
'                          one instance for each property retrieved.
'                          See the documentation in this procedure for
'                          details about calling it.
'       GetDesktopHandle - This function returns the handle of the Windows desktop.
'       GetProperty      - This procedure gets the value of the specified
'                          property.
'       RemoveProperty   - This procedure removes the property from the window's
'                          property list.
'       SetProperty      - This creates an new property or updates an existing
'                          property.
'       GetHWndOfForm    - This returns the HWnd of the UserForm that is passed
'                          in to the procedure. This is to be used if you are
'                          storing values in the UserForm window's property list.
'       GetNewCPropType  - This returns a New CPropType class instance. This
'                          procedure is intended to be used when calling these
'                          procedures for other VBProjects that reference this
'                          Project. If you import this module and the CPropType
'                          class into your project, you can create a new CPropType
'                          instance with the New keyword -- you don't need to
'                          use the GetNewCPropType function.
'
'       All of these procedures have an optional argument name HWnd. If this
'       argument is omitted or is <= 0, the properties are stored in the main
'       Excel application window's property list. If HWnd is included and is > 0,
'       the property  list for that window is used. If you want to store properties
'       in a UserForm's property list, you can call HWnd = GetHWndOfForm(UF:=YourFormName)
'       to retrieve the HWnd of the form, and pass that value in the HWnd parameter
'       to the various functions to set or retrieve the value.
'
' The following are the Private procedures that are used to support the Public
' procedures in this module. You don't need to access these Private procedures (that
' is why they are declared as Private). They are used to support the Public procedures.
'
'       IsArrayAllocated
'       IsArrayDynamic
'       IsArrayEmpty
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    
Private Declare Function IsWindow Lib "user32" ( _
    ByVal HWnd As Long) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" ( _
    ByVal HWnd As Long, _
    ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32" Alias "SetPropA" ( _
    ByVal HWnd As Long, _
    ByVal lpString As String, _
    ByVal hData As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" ( _
    ByVal HWnd As Long, _
    ByVal lpString As String) As Long

Private Declare Function EnumProps Lib "user32.dll" Alias "EnumPropsA" ( _
    ByVal HWnd As Long, _
    ByVal lpEnumFunc As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long


''''''''''''''''''''''''''''''''''''''''''''''''
' Note: The Visual Studio 6 API Viewer program
' shows the lpString type as String, not Long.
' It is incorrect. lpString needs to be a Long.
''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function LStrLen Lib "kernel32" Alias "lstrlenA" ( _
    ByVal lpString As Long) As Long
    
''''''''''''''''''''''''''''''''''''''''''''''''
' Note: The Visual Studio 6 API Viewer program
' shows the lpString2 type as String, not Long.
' It is incorrect.  lpString2 needs to be a Long.
''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function LStrCpy Lib "kernel32.dll" Alias "lstrcpyA" ( _
      ByVal lpString1 As String, _
      ByVal lpString2 As Long) As Long
    
''''''''''''''''''''''''''''''''''''''''
' These two variables are used with the
' GetAllProperties procedure. See the
' documentation in GetAllProperties
' for details.
''''''''''''''''''''''''''''''''''''''''
Private ArrayNdx As Long
Private ListAllArray() As CPropType


''''''''''''''''''''''''''''''''''''''''
' These two variables are used with the
' PropertyExists procedure. See the
' documentation in PropertyExists
' procedure for details.
''''''''''''''''''''''''''''''''''''''''
Private PropertyToFind As String
Private PropertyFound As Boolean

Public Function GetDesktopHandle() As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetDesktopHandle
' This returns the windows handle of the desktop window.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    GetDesktopHandle = GetDesktopWindow()
End Function



Public Function GetNewCPropType() As CPropType
''''''''''''''''''''''''''''''''''''''''''''''''
' GetNewCPropType
' This returns a new instance of CPropType to the
' calling procedure. This is to be used when you are
' calling these procedures from another VBAProject
' that references this project. If you import this
' module and the CPropType class into your project,
' you can simply create a new class instance with
' the New keyword. E.g.,
'
'     Dim PT As CPropType
'     Set PT = New CPropType
'
' The Instancing property of CPropType is
' PublicNotCreatable, so another project can
' declare a variable of that type, but not create
' an instance of the class. This function creates
' and returns a new instance of CPropType. E.g.,
'
'     Dim PT As projGetSetProps.CPropType
'     Set PT = projGetSetProps.GetNewCPropType()
'
''''''''''''''''''''''''''''''''''''''''''''''''
    Set GetNewCPropType = New CPropType
End Function


Public Function SetProperty(PropertyName As String, PropertyValue As Long, _
        Optional HWnd As Long = 0) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetProperty
' This function adds a property entry named PropertyName with the value
' PropertyValue to the window indentified by HWnd. If HWnd is omitted or
' <= 0, it is added to the main Excel application window's property list.
' The function returns True if the operation was successful, or False
' if an error occurred.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Res As Long
Dim DestHWnd As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If HWnd was omitted or <= 0, use the Excel main application
' window HWnd.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If HWnd <= 0 Then
    DestHWnd = FindWindow("XLMAIN", Application.Caption)
Else
    DestHWnd = HWnd
End If

If DestHWnd = 0 Then
    SetProperty = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' Ensure PropertyName is not an empty string.
'''''''''''''''''''''''''''''''''''''''''''''
If Trim(PropertyName) = vbNullString Then
    SetProperty = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' Ensure DestHWnd is an existing window.
'''''''''''''''''''''''''''''''''''''''''''''
If IsWindow(DestHWnd) = 0 Then
    SetProperty = False
    Exit Function
End If

Res = SetProp(HWnd:=DestHWnd, lpString:=PropertyName, hData:=PropertyValue)
If Res = 0 Then
    '''''''''''''''''''''
    ' An error occurred.
    '''''''''''''''''''''
    SetProperty = False
Else
    '''''''''''''''''''''
    ' Success.
    '''''''''''''''''''''
    SetProperty = True
End If

End Function

Public Function GetAllProperties(ResultArray As Variant, _
    Optional HWnd As Long = 0) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetAllProperties
' This procedure creates an array in ResultArray, each element of which
' is an instance of the CPropType class, containing the name and value
' of each property in the property list of the window specified by HWnd.
' If HWnd is omitted or <= 0, the main Excel application window's property
' list is used.
'
' ResultArray must be a dynamic, single-dimensional array. The existing
' contents of ResultArray will be destroyed.
'
' The function returns the number of elements added to ResultArray,
' or -1 if an error occurred. The calling procedure should declare
' a dynamic array of CPropType classes, each of which will store the
' name and value of a property:
'
'        Dim PropArray() As CPropType
'
' It should then pass that array to this procedure:
'
'        Dim Res As Long
'        Res = GetAllProperties(ResultArray:=PropArray, HWnd:=0)
'
' This procedure will Erase and then repopulate ResultArray with instances
' of CPropType objects. Upon return from this procedure, the calling
' procedure should then loop through the array:
'        If Res > 0 Then
'            For N = LBound(PropArray) To UBound(PropArray)
'                Debug.Print CStr(N), PropArray(N).Name, PropArray(N).Value
'            Next N
'        End If
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Res As Long
Dim DestHWnd As Long
Dim Counter As Long
Dim Ndx As Long
Dim PT As CPropType

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If HWnd was omitted or <= 0, use the Excel main application
' window HWnd.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If HWnd <= 0 Then
    DestHWnd = FindWindow("XLMAIN", Application.Caption)
Else
    DestHWnd = HWnd
End If

If DestHWnd = 0 Then
    GetAllProperties = -1
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure ResultArray is an array.
''''''''''''''''''''''''''''''''''
If IsArray(ResultArray) = False Then
    GetAllProperties = -1
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure ResultArray is dynamic.
''''''''''''''''''''''''''''''''''
If IsArrayDynamic(Arr:=ResultArray) = False Then
    GetAllProperties = -1
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' Ensure DestHWnd is an existing window.
'''''''''''''''''''''''''''''''''''''''''''''
If IsWindow(DestHWnd) = 0 Then
    GetAllProperties = -1
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' Erase the existing ListAllArray and set the
' ArrayNdx variable to 0.
''''''''''''''''''''''''''''''''''''''''''''''
Erase ListAllArray
Erase ResultArray
ArrayNdx = 0
'''''''''''''''''''''''''''''''''''''''''''''''
' Call EnumProps to get all the properties of
' DestHWnd's property list. Windows will call
' ProcEnumPropForListAll for each property
' in the window's property list.
'''''''''''''''''''''''''''''''''''''''''''''''
Res = EnumProps(HWnd:=DestHWnd, lpEnumFunc:=AddressOf ProcEnumPropForListAll)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Redim the ResultArray to the number of properties
' enumerated by EnumProps. Copy the array ListAllArray
' to ResultArray.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(Arr:=ListAllArray) = True Then
    ReDim ResultArray(1 To UBound(ListAllArray))
    Set PT = New CPropType
    For Ndx = LBound(ListAllArray) To UBound(ListAllArray)
        Set PT = ListAllArray(Ndx)
        PT.Name = ListAllArray(Ndx).Name
        PT.Value = ListAllArray(Ndx).Value
        Set ResultArray(Ndx) = PT
    Next Ndx
End If
''''''''''''''''''''''''''''''''''''''''''''''''''
' If the array is allocated, we retrieved at least
' one property. Return the number of properties
' retrieved. If the array is not allocated, there
' were no properties to retrieve, so return 0.
''''''''''''''''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(Arr:=ResultArray) = True Then
    GetAllProperties = UBound(ResultArray)
Else
    GetAllProperties = 0
End If


End Function


Public Function GetProperty(PropertyName As String, ByRef PropertyValue As Long, _
    Optional HWnd As Long = 0) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetProperty
' This function retrieves the value of PropertyName from
' the window specified by HWnd. If HWnd is omitted or <= 0,
' it looks in the main Excel application window's property
' list. It will place the value of the specified property
' in the variable passed as PropertyValue. You must pass
' a Long type of variable for PropertyValue.
' The function returns True if the operation was successful,
' or False if an error occurred. If an error occurs, the
' variable PropertyValue is left unchanged.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Res As Long
Dim DestHWnd As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If HWnd was omitted or is <= 0, use the Excel main application
' window HWnd.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If HWnd <= 0 Then
    DestHWnd = FindWindow("XLMAIN", Application.Caption)
Else
    DestHWnd = HWnd
End If

'''''''''''''''''''''''''''''''''''''''''''''
' Ensure DestHWnd is an existing window.
'''''''''''''''''''''''''''''''''''''''''''''
If IsWindow(DestHWnd) = 0 Then
    GetProperty = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' Ensure PropertyName is not an empty string.
'''''''''''''''''''''''''''''''''''''''''''''
If Trim(PropertyName) = vbNullString Then
    GetProperty = False
    Exit Function
End If

Res = GetProp(DestHWnd, PropertyName)
'''''''''''''''''''''''''''''''''''''
' GetProp will return 0 if an error
' occurred, but 0 may also be a valid
' property value. Test Err.LastDllError
' to see if an error occurred. If it
' indicates an error, it is most likely
' that the property doesn't exist
' (Err.LastDllError = 2).
'''''''''''''''''''''''''''''''''''''
If Err.LastDllError <> 0 Then
    GetProperty = False
Else
    PropertyValue = Res
    GetProperty = True
End If

End Function



Public Function PropertyExists(PropertyName As String, _
    Optional HWnd As Long = 0) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PropertyExists
' This function returns True or False indicating whether the
' property with the string value PropertyName exists for the
' window specified in HWnd. If HWnd is omitted or <= 0, the
' main Excel application window's property list is searched.
' The function returns True if the property exists or False
' if the property does not exist or an error occurred.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Res As Long
Dim DestHWnd As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If HWnd was omitted or <= 0, use the Excel main application
' window HWnd.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If HWnd <= 0 Then
    DestHWnd = FindWindow("XLMAIN", Application.Caption)
Else
    DestHWnd = HWnd
End If


'''''''''''''''''''''''''''''''''''''''''''''
' Ensure DestHWnd is an existing window.
'''''''''''''''''''''''''''''''''''''''''''''
If IsWindow(DestHWnd) = 0 Then
    PropertyExists = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''
' Set PropetyFound to False and set PropertyToFind
' the the property name we're looking for.
''''''''''''''''''''''''''''''''''''''''''''''''''
PropertyFound = False
PropertyToFind = PropertyName
Res = EnumProps(DestHWnd, AddressOf ProcEnumPropForFind)

If PropertyFound = True Then
    PropertyExists = True
Else
    PropertyExists = False
End If

End Function


Public Function RemoveProperty(PropertyName As String, _
    Optional HWnd As Long = 0) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RemoveProperty
' This function removes the property named by PropertyName from the property
' list of the window specified by HWnd. If HWnd is omitted or <= 0, then
' main Excel application window's property list is used.
' The function returns True if the operation was successful, or False if
' an error occurred.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Res As Long
Dim DestHWnd As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If HWnd was omitted or <= 0, use the Excel main application
' window HWnd.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If HWnd <= 0 Then
    DestHWnd = FindWindow("XLMAIN", Application.Caption)
Else
    DestHWnd = HWnd
End If

'''''''''''''''''''''''''''''''''''''''''''''
' Ensure DestHWnd is an existing window.
'''''''''''''''''''''''''''''''''''''''''''''
If IsWindow(DestHWnd) = 0 Then
    RemoveProperty = False
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''''
' Ensure PropertyName is not an empty string.
'''''''''''''''''''''''''''''''''''''''''''''
If Trim(PropertyName) = vbNullString Then
    RemoveProperty = False
    Exit Function
End If

Res = RemoveProp(DestHWnd, PropertyName)
''''''''''''''''''''''''''''''''
' If PropertyName doesn't exist
' we'll get an error value in Res.
' We can safely ignore this error
' and return True.
''''''''''''''''''''''''''''''''
RemoveProperty = True
End Function

Public Function GetHWndOfForm(UF As Object) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetHWndOfForm
' This returns the HWnd of the UserForm referenced in UF.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim HWnd As Long
    HWnd = FindWindow("ThunderDFrame", UF.Caption)
    GetHWndOfForm = HWnd
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private Support Procedures
' These functions are documented and available for download at
' http://www.cpearson.com/excel/vbaarrays.htm.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IsArrayAllocated(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayAllocated
' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
' allocated.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is just the reverse of IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim N As Long
On Error Resume Next

' if Arr is not an array, return FALSE and get out.
If IsArray(Arr) = False Then
    IsArrayAllocated = False
    Exit Function
End If

' Attempt to get the UBound of the array. If the array has not been allocated,
' an error will occur. Test Err.Number to see if an error occurred.
N = UBound(Arr, 1)
If (Err.Number = 0) Then
    ''''''''''''''''''''''''''''''''''''''
    ' Under some circumstances, if an array
    ' is not allocated, Err.Number will be
    ' 0. To acccomodate this case, we test
    ' whether LBound <= Ubound. If this
    ' is True, the array is allocated. Otherwise,
    ' the array is not allocated.
    '''''''''''''''''''''''''''''''''''''''
    If LBound(Arr) <= UBound(Arr) Then
        ' no error. array has been allocated.
        IsArrayAllocated = True
    Else
        IsArrayAllocated = False
    End If
Else
    ' error. unallocated array
    IsArrayAllocated = False
End If

End Function

Private Function IsArrayDynamic(ByRef Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayDynamic
' This function returns TRUE or FALSE indicating whether Arr is a dynamic array.
' Note that if you attempt to ReDim a static array in the same procedure in which it is
' declared, you'll get a compiler error and your code won't run at all.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim LUBound As Long

' If we weren't passed an array, get out now with a FALSE result
If IsArray(Arr) = False Then
    IsArrayDynamic = False
    Exit Function
End If

' If the array is empty, it hasn't been allocated yet, so we know
' it must be a dynamic array.
If IsArrayEmpty(Arr:=Arr) = True Then
    IsArrayDynamic = True
    Exit Function
End If

' Save the UBound of Arr.
' This value will be used to restore the original UBound if Arr
' is a single-dimensional dynamic array. Unused if Arr is multi-dimensional,
' or if Arr is a static array.
LUBound = UBound(Arr)

On Error Resume Next
Err.Clear

' Attempt to increase the UBound of Arr and test the value of Err.Number.
' If Arr is a static array, either single- or multi-dimensional, we'll get a
' C_ERR_ARRAY_IS_FIXED_OR_LOCKED error. In this case, return FALSE.
'
' If Arr is a single-dimensional dynamic array, we'll get C_ERR_NO_ERROR error.
'
' If Arr is a multi-dimensional dynamic array, we'll get a
' C_ERR_SUBSCRIPT_OUT_OF_RANGE error.
'
' For either C_NO_ERROR or C_ERR_SUBSCRIPT_OUT_OF_RANGE, return TRUE.
' For C_ERR_ARRAY_IS_FIXED_OR_LOCKED, return FALSE.

ReDim Preserve Arr(LBound(Arr) To LUBound + 1)

Select Case Err.Number
    Case 0
        ' We successfully increased the UBound of Arr.
        ' Do a ReDim Preserve to restore the original UBound.
        ReDim Preserve Arr(LBound(Arr) To LUBound)
        IsArrayDynamic = True
    Case 9
        ' Arr is a multi-dimensional dynamic array.
        ' Return True.
        IsArrayDynamic = True
    Case 10
        ' Arr is a static single- or multi-dimensional array.
        ' Return False
        IsArrayDynamic = False
    Case Else
        ' We should never get here.
        ' Some unexpected error occurred. Be safe and return False.
        IsArrayDynamic = False
End Select

End Function


Private Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayEmpty
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is really the reverse of IsArrayAllocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Var As Variant
Err.Clear
On Error Resume Next
If IsArray(Arr) = False Then
    ' we weren't passed an array, return True
    IsArrayEmpty = True
End If

' Attempt to get the UBound of the array. If the array is
' unallocated, an error will occur.
Var = UBound(Arr, 1)
If (Err.Number <> 0) Or (Var < 0) Then
    IsArrayEmpty = True
Else
    IsArrayEmpty = False
End If

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows Callback procedures for EnumProps
' These addresses of these procedures are passed to the EnumProps API function.
' Windows will call the procedure passed to EnumProps one time for each property
' in the specified window's property list. These procedures MUST be declared
' exactly as shown. If you change the declaration, you'll crash Excel.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function ProcEnumPropForFind(ByVal HWnd As Long, ByVal Addr As Long, _
            ByVal Data As Long) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ProcEnumPropForFind
' This is the Windows callback function for determining if a property exits.  It
' is called by Windows for each property in the property list. We test the string
' value provided to this procedure against the value of PropertyToFind. If we get
' a match, the property exists and the PropertyFound value is set to True, and
' we terminate the enumeration.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim StringName As String
Dim Res As Long
Dim SLen As Long
Dim Pos As Integer

'''''''''''''''''''''''''''''''
' Set the PropertyFound variable
' to False.
''''''''''''''''''''''''''''''''
PropertyFound = False

'''''''''''''''''''''''''''''''
' Get the length of the string
' stored at the address Addr.
' This length does not include
' the trailing null character.
'''''''''''''''''''''''''''''''
SLen = LStrLen(Addr)
'''''''''''''''''''''''''''''''
' Allocate the StringName buffer.
' The +1 is to make room for the
' trailing null character.
'''''''''''''''''''''''''''''''
StringName = String$(SLen + 1, vbNullChar)
'''''''''''''''''''''''''''''''''''
' Copy the string from Addr to the
' StringName buffer variable.
'''''''''''''''''''''''''''''''''''
Res = LStrCpy(ByVal StringName, Addr)
If Res = 0 Then
    Debug.Print "An error occurred with LStrCpy.", Err.LastDllError
Else
    '''''''''''''''''''''''''''''''''''''''
    ' Trim off the trailing null character.
    '''''''''''''''''''''''''''''''''''''''
    Pos = InStr(1, StringName, vbNullChar)
    If Pos > 0 Then
        StringName = Left(StringName, Pos - 1)
    End If
    ''''''''''''''''''''''''''''''''''''''
    ' Compare PropertyName to StringName.
    ' If they match, set PropertyFound
    ' to True and terminate the enumeration.
    ''''''''''''''''''''''''''''''''''''''
    If StrComp(PropertyToFind, StringName, vbTextCompare) = 0 Then
        PropertyFound = True
        ProcEnumPropForFind = False
        Exit Function
    End If
End If
'''''''''''''''''''''''''''''
' Return True to continue the
' enumeration.
'''''''''''''''''''''''''''''
ProcEnumPropForFind = True

End Function

Private Function ProcEnumProp(ByVal HWnd As Long, ByVal Addr As Long, _
            ByVal Data As Long) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ProcEnumProp
' This is the callback function for EnumProps. Windows will call
' this function for each Property associated with the HWnd in the
' call to EnumProps.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim StringName As String
Dim Res As Long
Dim SLen As Long
Dim Pos As Integer

'''''''''''''''''''''''''''''''
' Get the length of the string
' stored at the address Addr.
' This length does not include
' the trailing null character.
'''''''''''''''''''''''''''''''
SLen = LStrLen(Addr)
'''''''''''''''''''''''''''''''
' Allocate the StringName buffer.
' The +1 is to make room for the
' trailing null character.
'''''''''''''''''''''''''''''''
StringName = String$(SLen + 1, vbNullChar)
'''''''''''''''''''''''''''''''''''
' Copy the string from Addr to the
' StringName buffer variable.
'''''''''''''''''''''''''''''''''''
Res = LStrCpy(ByVal StringName, Addr)
If Res = 0 Then
    Debug.Print "An error occurred with LStrCpy.", Err.LastDllError
Else
    '''''''''''''''''''''''''''''''''''''''
    ' Trim off the trailing null character.
    '''''''''''''''''''''''''''''''''''''''
    Pos = InStr(1, StringName, vbNullChar)
    If Pos > 0 Then
        StringName = Left(StringName, Pos - 1)
    End If
    Debug.Print CStr(Addr), StringName, CStr(Data)
End If
'''''''''''''''''''''''''''''
' Return True to continue the
' enumeration.
'''''''''''''''''''''''''''''
ProcEnumProp = True
End Function


Private Function ProcEnumPropForListAll(ByVal HWnd As Long, ByVal Addr As Long, _
            ByVal Data As Long) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ProcEnumPropForListAll
' This is the Windows callback procedure for EnumProps called by GetAllProperties. It
' stores each property name and associated value in a CPropType class instance and
' adds that to the module-level variable ListAllArray. ListAllArray should be Erased
' and ArrayNdx set to 0 prior to calling the EnumProps API function that calls this
' function.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim StringName As String
Dim Res As Long
Dim SLen As Long
Dim Pos As Integer
Dim PropType As CPropType
'''''''''''''''''''''''''''''''''
' Get the length of the string.
' This length does not include
' the trailing null character.
'''''''''''''''''''''''''''''''''
SLen = LStrLen(Addr)
'''''''''''''''''''''''''''''''''
' Allocate StringName to SLen+1
' vbNullChars. The +1 is for the
' trailing null character.
'''''''''''''''''''''''''''''''''
StringName = String$(SLen + 1, vbNullChar)
'''''''''''''''''''''''''''''''''''''''
' Copy the string from the address Addr
' to the StringName buffer variable.
'''''''''''''''''''''''''''''''''''''''
Res = LStrCpy(ByVal StringName, Addr)
''''''''''''''''''''''''''''''''''''''
' Trim to the vbNullChar if necessary.
''''''''''''''''''''''''''''''''''''''
Pos = InStr(1, StringName, vbNullChar)
If Pos > 0 Then
    StringName = Left(StringName, Pos - 1)
End If
'''''''''''''''''''''''''''''''''''''''''
' Create a new instance of CPropType,
' increment the array index and resize
' the array. Set the last element of
' the array to the new CPropType variable.
'''''''''''''''''''''''''''''''''''''''''
Set PropType = New CPropType
ArrayNdx = ArrayNdx + 1
ReDim Preserve ListAllArray(1 To ArrayNdx)
PropType.Name = StringName
PropType.Value = Data
Set ListAllArray(UBound(ListAllArray)) = PropType

'''''''''''''''''''''''''''''
' Return True to continue the
' enumeration.
'''''''''''''''''''''''''''''
ProcEnumPropForListAll = True
End Function



