VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "floatingHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const DomainErrNum As Long = vbObjectError + 1000 + 1
Const DomainErrDesc As String = "Columns and Rows Headers must be from the same Worksheet"

Const isDynamic As Boolean = False
Const useTimer As Boolean = True

Const DEFcolHdrRange As String = "colsHDRRange"
Const DEFrowHdrRange As String = "rowsHDRRange"
Const CHkey As String = "colHDR"
Const RHkey As String = "rowHDR"

Private mOnTime As cOnTime

Private WithEvents mWs As Worksheet
Private WithEvents mWb As Workbook
Attribute mWb.VB_VarHelpID = -1

Private mWindowLeft As Long
Private mRowsPicWidth As Long

Property Let columnsHdr(columnsHdrRange As Range)
Dim actCell As Range
    If setWorksheet(columnsHdrRange) Then
        Set actCell = ActiveCell
        rangePic columnsHdrRange, CHkey
        actCell.Activate
    End If
End Property

Property Let rowsHdr(rowsHdrRange As Range)
Dim actCell As Range
    If setWorksheet(rowsHdrRange) Then
        Set actCell = ActiveCell
        rangePic rowsHdrRange, RHkey
        mRowsPicWidth = rowsHdrRange.Columns(1).Width
        actCell.Activate
    End If
End Property

Private Sub rangePic(picRange As Range, picName As String)
Dim pic As Picture

    picRange.CopyPicture
    picRange.Cells(1).PasteSpecial
    With Selection
        If isDynamic Then
            On Error Resume Next
            .Formula = picRange.Address(ReferenceStyle:=Application.ReferenceStyle)
            On Error GoTo 0
            .Formula = "=" & picRange.Address(ReferenceStyle:=Application.ReferenceStyle)
        End If
        .Name = picName
        .Top = picRange.Cells(1).Top
        .Left = picRange.Cells(1).Left
    End With
    
End Sub
Property Get isActiveSheet() As Boolean
Dim ws As Object
    Set ws = ActiveSheet
    On Error GoTo exitfalse
    isActiveSheet = ws.isActiveSheet
    Exit Property
exitfalse:
    Err.Clear
    isActiveSheet = False
End Property
Private Function setWorksheet(r As Range) As Boolean
    
    If isActiveSheet Then
        If mWs Is Nothing Then
            Set mWs = r.Worksheet
            setWorksheet = True
        Else
            If r.Worksheet.CodeName <> mWs.CodeName Then
                Err.Raise DomainErrNum, "floatingHeaders Class", DomainErrDesc
            Else
            setWorksheet = True
            End If
        End If
    Else
        setWorksheet = False
    End If

End Function

Private Sub Class_Initialize()
Const myName As String = "floatingHeaders.Class_Initialize"
Dim ws As Worksheet
    
    Set ws = ActiveSheet
    Debug.Print myName & vbTab & ws.Name
    If isActiveSheet Then
        Debug.Print myName & vbTab & ws.Name & " is active sheet"
        clearPics ws
        columnsHdr = Range(DEFcolHdrRange)
        rowsHdr = Range(DEFrowHdrRange)
        Set mOnTime = New cOnTime
    End If
    Set mWb = ActiveWorkbook
End Sub

Private Sub Class_Terminate()
Const myName As String = "floatingHeaders.Class_Terminate"
    Debug.Print myName
    
    On Error Resume Next
    clearPics mWs
    Set mOnTime = Nothing

End Sub
Private Sub clearPics(ws As Worksheet)
Const myName As String = "floatingHeaders.clearPics"
Dim shp As Shape, actWS As Worksheet
    
    If ActiveSheet.CodeName <> ws.CodeName Then
        Application.ScreenUpdating = False
        Set actWS = ActiveSheet
        ws.Activate
    End If
    
    Debug.Print myName & ws.Name & " Shapes:"
    For Each shp In ws.Shapes
        Debug.Print vbTab & shp.Name
    Next shp
    Debug.Print vbTab & "Error state: " & Err.Number
    
    On Error Resume Next
    Do
        DoEvents
        ws.Shapes(CHkey).Delete
        DoEvents
'        If Err.Number = 0 Then Debug.Print vbTab & "Deleted " & CHkey & " in " & ws.Name & vbTab & ws.Shapes.Count & " remaining"
    Loop While Err.Number = 0
    Err.Clear
    Debug.Print vbTab & "Deleted " & CHkey & " in " & ws.Name & vbTab & ws.Shapes.Count & " remaining"
    
    Do
        DoEvents
        ws.Shapes(RHkey).Delete
        DoEvents
    Loop While Err.Number = 0
    Err.Clear
    Debug.Print vbTab & "Deleted " & RHkey & " in " & ws.Name & vbTab & ws.Shapes.Count & " remaining"
    
    On Error GoTo 0
    
    Debug.Print "Pics deleted " & ws.Shapes.Count & " pics remain"
    
    If Not actWS Is Nothing Then
        actWS.Activate
        Application.ScreenUpdating = True
    End If

End Sub
Private Sub alignWithWindow()
Dim wl As Long, vRangeLeftTop As Range
On Error Resume Next
    Set vRangeLeftTop = ActiveWindow.VisibleRange.Cells(1, 1)
    If vRangeLeftTop.Width < mRowsPicWidth Then
        Set vRangeLeftTop = vRangeLeftTop.Offset(0, 1)
    End If
    With vRangeLeftTop
        wl = .Left + .Width - mRowsPicWidth
    End With
    If wl <> mWindowLeft Then
        mWindowLeft = wl
        mWs.Shapes(CHkey).Left = wl
        mWs.Shapes(RHkey).Left = wl
    End If
End Sub
Private Sub mws_SelectionChange(ByVal Target As Range)
    If Not useTimer Then alignWithWindow
End Sub
Public Sub kickDog()
    mOnTime.kickDog
    alignWithWindow
End Sub
Private Sub mws_Activate()
Const myName As String = "Sheet2.Worksheet_Activate"
    Debug.Print myName

    If useTimer Then Set mOnTime = New cOnTime
    
End Sub

Private Sub mws_Deactivate()
Const myName As String = "Sheet2.Worksheet_Deactivate"
    Debug.Print myName

    On Error Resume Next
    Set mOnTime = Nothing
    
End Sub

Private Sub mWb_SheetActivate(ByVal Sh As Object)
Const myName As String = "floatingHeaders.mWb_SheetActivate"

    Debug.Print myName
    If mWs Is Nothing Then Class_Initialize

End Sub

