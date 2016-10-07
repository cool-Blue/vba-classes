VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassComboBoxEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'http://www.excelforum.com/excel-programming-vba-macros/1021112-call-userform-from-a-variable-number-of-activex-command-buttons.html

Private mOLE As OLEObject
Private WithEvents mCB As MSForms.ComboBox
Attribute mCB.VB_VarHelpID = -1
Private mWS As Worksheet
Public ListFound As Boolean
Public Quiet As Boolean
Private Sub Class_Initialize()
    Quiet = True
End Sub
Property Get OLE() As OLEObject
    Set OLE = mOLE
End Property
Property Let Output(RHS As Range)
Attribute Output.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Output.VB_UserMemId = 0
    OLE = RHS.Worksheet.OLEObjects.Add(ClassType:="Forms.ComboBox.1", _
                                        Link:=False, _
                                        DisplayAsIcon:=False, _
                                        Left:=RHS.Left, _
                                        Top:=RHS.Top, _
                                        Width:=RHS.Width * 2, _
                                        Height:=RHS.Height * 2)
End Property
Property Let OLE(RHS As OLEObject)
    Set mOLE = RHS
    Set mWS = RHS.Parent
    Select Case TypeName(RHS.Object)
    Case "ComboBox"
        Set mCB = RHS.Object
        On Error Resume Next
        mCB.List = cbListRange(mOLE).Value
        If Err.Number <> 0 Then
        '   Cant find a definition block
            ListFound = False
        Else
            ListFound = True
        End If
        Debug.Print mOLE.Name & " connected"
    Case Else
        MsgBox "Unsupported Control Type: " & TypeName(RHS.Object)
    End Select
End Property
Property Get Name() As String
    Name = mOLE.Name
End Property

Private Sub Class_Terminate()
Dim msg As String
    If mOLE Is Nothing Then
        msg = "mOLE is Nothing"
    Else
        On Error Resume Next
        mOLE.Activate
        msg = mOLE.Name
        If Err.Number <> 0 Then
            msg = Err.Description
        End If
    End If
    Debug.Print "Terminating " & msg
End Sub

Private Sub mCB_Change()
  'This is the Active X Event Handler
  If Not Quiet Then MsgBox "mCB_Change():  " & mCB.Name & "     Value = " & mCB.Value
End Sub

Private Function cbListRange(oOLE As OLEObject) As Range
' Find the header based on the name of the OLE object and return
' the list.

  Dim r As Range, sFindString As String
  
  sFindString = oOLE.Name & " Items"
  
  'Find the first occurence of the string
  Set r = Nothing
  Set r = mWS.Cells.Find(What:=sFindString, _
                      After:=mWS.Range("A1"), _
                      LookIn:=xlValues, _
                      LookAt:=xlWhole, _
                      SearchOrder:=xlByRows, _
                      SearchDirection:=xlNext, _
                      MatchCase:=False, _
                      SearchFormat:=False)
                      
                      
  If Not r Is Nothing Then
  ' Return the list of contiguous cells, starting from the next row below the header
    Set cbListRange = r
    Set cbListRange = r.Offset(1)
    Set cbListRange = mWS.Range(cbListRange, cbListRange.End(xlDown))
  End If
 
End Function


