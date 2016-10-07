VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mCollection As Collection

Private Sub Class_Initialize()
    Set mCollection = New Collection
End Sub
Private Sub Class_Terminate()
    Set mCollection = Nothing
End Sub

Public Sub AddArray(ParamArray arr() As Variant)
Dim v As Variant
    For Each v In arr
        mCollection.Add v
    Next
End Sub

Public Sub Add(item As Variant, Optional key As String = vbNullString)
  If key = vbNullString Then
    If TypeName(item) = "Object" Then
      On Error Resume Next
      key = item.Name
      On Error GoTo 0
    End If
  End If
  If key = vbNullString Then
    mCollection.Add item
  Else
    mCollection.Add item, key
  End If
End Sub

Public Sub Remove(index As Variant)
    mCollection.Remove index
End Sub

Public Property Get item(index As Variant) As Variant
Attribute item.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Item.VB_UserMemId = 0
  If Not isIncluded(CStr(index)) Then
    item = Null
    Exit Property
  End If
  If IsObject(mCollection.item(index)) Then
    Set item = mCollection.item(index)
  Else
    item = mCollection.item(index)
  End If
End Property
Public Property Get ItemOrNull(index As Variant) As Variant

  If IsNumeric(index) Then
    If index > mCollection.Count Or index < 0 Then
      ItemOrNull = Null
      Exit Property
    End If
  Else
    If Not isIncluded(CStr(index)) Then
      ItemOrNull = Null
      Exit Property
    End If
  End If
  
  If IsObject(ItemOrNull) Then
    Set ItemOrNull = mCollection.item(index)
  Else
    ItemOrNull = mCollection.item(index)
  End If
  
End Property
Public Property Get Count() As Long
    Count = mCollection.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Enumerator: Edit in Notepad to see this.
'    Attribute NewEnum.VB_UserMemId = -4
'    Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCollection.[_NewEnum]
End Property

Public Function isIncluded(key As String) As Boolean
Const NOT_INCLUDED As Long = 5
    On Error GoTo incol
    mCollection.item key

incol:
    isIncluded = (Err.Number <> NOT_INCLUDED)
    If Err.Number = NOT_INCLUDED Then Err.Clear
End Function

Public Sub dbReport()
  mdbReport mCollection
End Sub
Private Sub mdbReport(c As Collection)
Dim db As New cDebugReporter

Dim v As Variant, c1 As Collection

  For Each v In c
    If TypeName(v) = "Collection" Then
      Set c1 = v
      mdbReport c1
    Else
      If Not IsObject(v) Then
        db.Report message:=CStr(v)
      Else
        db.Report message:="Object: " & TypeName(v)
      End If
    End If
  Next v

End Sub

