VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cldefenSheetObjects_NameSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcrb As clBordersRegion
Private mWks As Worksheet
Private mSpaceName As String
Private mSpaceType As String
Private mSpaceRange As Range
Private mHeaderRange As Range
Private mVariables As Collection
Public Sub nameSpace(wks As Worksheet, spaceType As String, Name As String)
    Set mWks = wks
    mSpaceType = spaceType
    mSpaceName = Name
End Sub
Public Function addVariableUnTyped(varName As String) As Boolean
Dim v As cldefSheetObjects_Variable, nameRange As Range
    
    If mVariables Is Nothing Then
        Set mVariables = New Collection
    Else
        If isIncluded(mVariables, varName) Then GoTo exitFail
    End If
    
    Set v = createVariable(varName)
    If v Is Nothing Then GoTo exitFail
    
    mVariables.Add v, varName
    addVariableUnTyped = True
    Exit Function
exitFail:
    addVariableUnTyped = False
End Function
Public Function addVariableTyped(varName As String, varType As String) As Boolean
Dim v As cldefSheetObjects_Variable, nameRange As Range

    If mVariables Is Nothing Then
        Set mVariables = New Collection
    Else
        If isIncluded(mVariables, varName) Then GoTo exitFail
    End If
    Set v = createVariable(varName, varType)
    If v Is Nothing Then GoTo exitFail
    
    mVariables.Add v, varName
    addVariableTyped = True
    Exit Function
exitFail:
    addVariableTyped = False
End Function
Private Function createVariable(varName As String, Optional varType As String = "") As cldefSheetObjects_Variable
Dim v As cldefSheetObjects_Variable, nameRange As Range

    Set nameRange = varNameRange(varName)
    If nameRange Is Nothing Then GoTo exitFail
    
    Set v = New cldefSheetObjects_Variable
    With v
        .nameRange = nameRange
        If Len(varType) <> 0 Then
            If UCase(Trim(varType)) <> UCase(Trim(.varType)) Then GoTo exitFail
        End If
    End With 'v
    Set createVariable = v
    Exit Function
exitFail:
    Set createVariable = Nothing
End Function
Property Get Variable(varName As String) As cldefSheetObjects_Variable
Attribute Variable.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Variable.VB_UserMemId = 0
    If isIncluded(mVariables, varName) Then
        Set Variable = mVariables.Item(varName)
    Else
        If addVariableUnTyped(varName) Then
            Set Variable = mVariables.Item(varName)
        End If
    End If
End Property
Property Get Name() As String
    Name = mSpaceName
End Property
Property Get SpaceRange() As Range
    If mSpaceRange Is Nothing Then _
        Set mSpaceRange = getSpaceRange(mSpaceType, mSpaceName)
    Set SpaceRange = mSpaceRange
End Property
Property Get Header() As Range
    If mHeaderRange Is Nothing Then _
        Set mHeaderRange = getHeaderRange(mWks, mSpaceType, mSpaceName)
    Set Header = mHeaderRange
End Property
Private Function getSpaceRange(spaceType As String, Name As String) As Range
Dim mRange As Range
    If (Len(spaceType) = 0) Or (Len(Name) = 0) Then Exit Function
    Set mRange = Header
    If mRange Is Nothing Then GoTo exitNothing
    Set mRange = mRange.Resize(Rowsize:=mcrb.currentRegionBorders(mRange.Cells(1), xlDown).Rows.Count)
exitNothing:
    Set getSpaceRange = mRange
End Function
Private Function getHeaderRange(wks As Worksheet, spaceType As String, Name As String) As Range
Dim pTitle As Range, pPrevious As Range
    Set pPrevious = wks.Cells(1, 1)
    Set pTitle = wks.Cells.Find(Name, pPrevious)
    Do While Trim(pTitle.Offset(0, 1)) <> spaceType And Intersect(pTitle, pPrevious) Is Nothing
        Set pTitle = wks.Cells.FindNext(pTitle)
    Loop
    If Trim(pTitle.Offset(0, 1)) <> spaceType Then
        Set getHeaderRange = Nothing
    Else
        Set getHeaderRange = mcrb.currentRegionBorders(pTitle, xlToRight)
    End If
End Function

Property Get varNameRange(varName As String) As Range
    Set varNameRange = SpaceRange.Find(varName)
End Property
Private Function isIncluded(col As Variant, key As String) As Boolean
Const NOT_INCLUDED As Long = 5
    On Error GoTo incol
    col.Item key
incol:
    isIncluded = Not (col Is Nothing) And (Err.Number <> NOT_INCLUDED)
End Function

Private Sub Class_Initialize()
    Set mcrb = New clBordersRegion
End Sub

Private Sub Class_Terminate()
    On Error Resume Next     'allow for Collection is already Nothing
    Set mVariables = Nothing
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Enumerator: Edit in Notepad to see this.
'    Attribute NewEnum.VB_UserMemId = -4
'    Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mVariables.[_NewEnum]
End Property

