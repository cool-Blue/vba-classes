VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cfloatHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/ © cool.blue@y7mail.com
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Const RangeErrNum As Long = vbObjectError + 1000 + 1
Const RangeErrDesc As String = "Range must be on Active Sheet"
Const ConfigErrNum As Long = vbObjectError + 1000 + 2

Const isDynamic As Boolean = False

Const DEFcolHdrRange As String = "colsHDRRange"
Const DEFrowHdrRange As String = "rowsHDRRange"
Const CHkey As String = "colHDR"
Const RHkey As String = "rowHDR"

Const DEFisActiveSheet As String = "floatingHeaders"
Const DEFPulseTime = "PulseTime"
Const DEFearliestTime As Long = 500

Private WithEvents mWinTimer  As cWinTimer

Private WithEvents mHostSheet As cActiveSheet
Attribute mHostSheet.VB_VarHelpID = -1
Private mActivatedSheets As Collection

Private mWindowLeft As Long
Private mRowsPicWidth As Long
Private mScrollColumn As Long

Private Property Get mWS() As Worksheet
    On Error Resume Next
    Set mWS = mHostSheet.ws
    On Error GoTo 0
End Property

Private Function sheetActivated(shName As String) As Boolean
Dim pName As String

    On Error GoTo exitFalse
    pName = mActivatedSheets(shName).name
    sheetActivated = True
    Exit Function
    
exitFalse:
    sheetActivated = False
End Function
Private Sub addActiveSheet(sh As Worksheet)
    On Error Resume Next
    mActivatedSheets.Add sh, sh.CodeName
    On Error GoTo 0
End Sub
Private Function picExists(picName As String) As Boolean
Dim pName As String

    On Error GoTo exitFalse
    pName = mHostSheet.ws.Shapes(picName).name
    picExists = True
    Exit Function
    
exitFalse:
    picExists = False
End Function
Private Function rangePic(picRangeName As String, picName As String) As Long
Dim actCell As Range, scrnUpdate As Boolean
Dim picRange As Range

    If Not picExists(picName) Then
        
        With Application
            scrnUpdate = .ScreenUpdating
            If .ScreenUpdating Then .ScreenUpdating = False
        End With
        
        Set actCell = ActiveCell
        
        On Error GoTo rangeError
        Set picRange = Range(picRangeName)
        If picRange.Worksheet.CodeName <> mHostSheet.ws.CodeName Then GoTo rangeError
        On Error GoTo 0
        
        picRange.CopyPicture
        picRange.Cells(1).PasteSpecial
        With Selection
            If isDynamic Then
                On Error Resume Next
                .Formula = picRange.Address(ReferenceStyle:=Application.ReferenceStyle)
                On Error GoTo 0
                .Formula = "=" & picRange.Address(ReferenceStyle:=Application.ReferenceStyle)
            End If
            .name = picName
            .Top = picRange.Cells(1).Top
            .Left = picRange.Cells(1).Left
        End With
        
cleanUp:
        actCell.Activate
        With Application
            If scrnUpdate Then .ScreenUpdating = True
        End With
        
    End If
    
    rangePic = Range(picRangeName).Columns(1).Width
    
    Exit Function
    
rangeError:
    If Err.number <> 0 Then Err.Clear
    Err.Raise RangeErrNum, , RangeErrDesc
    Resume cleanUp

End Function
Private Sub clearPics(ws As Worksheet)
Const myName As String = "cfloatHeaders.clearPics"
Dim db As New cDebugReporter
    db.Report caller:=myName
    
    On Error Resume Next
    Do
        ws.Shapes(CHkey).Delete
    Loop While Err.number = 0
    Err.Clear
    
    Do
        ws.Shapes(RHkey).Delete
    Loop While Err.number = 0
    Err.Clear
    
    On Error GoTo 0
    
    db.ExitMessage = " Pics deleted " & ws.Shapes.Count & " pics remain"
    
End Sub
Private Sub alignWithWindow(Optional reSet As Boolean = False)
Dim wl As Double, vRangeLeftTop As Range
Dim shpRange As ShapeRange, i As Long

    If ActiveWindow.ScrollColumn <> mScrollColumn Or reSet Then
        mScrollColumn = ActiveWindow.ScrollColumn
        
        Set vRangeLeftTop = ActiveWindow.VisibleRange.Cells(1, 1)
        If vRangeLeftTop.Width < mRowsPicWidth Then
            Set vRangeLeftTop = vRangeLeftTop.Offset(0, 1)
        End If
        With vRangeLeftTop
            i = 0
            Do
                wl = .Offset(0, i).Left + .Offset(0, i).Width - mRowsPicWidth
                i = i + 1
            Loop Until wl >= .Left
        End With
        If wl <> mWindowLeft Or reSet Then
            mWindowLeft = wl
            Set shpRange = mWS.Shapes.Range(Array(CHkey, RHkey))
            shpRange.Left = wl
        End If
        
    End If

End Sub
Private Function pulseTime() As Long
Const myName As String = "cfloatHeaders.pulseTime"

    On Error GoTo propertyFailed
    pulseTime = mHostSheet.wsGet(DEFPulseTime)
    Exit Function
propertyFailed:
    If mHostSheet.isQualified Then
        pulseTime = DEFearliestTime
    Else
        pulseTime = -1
    End If
    
End Function
Public Sub startTimer(pulseTime As Long)
Const myName As String = "cfloatHeaders.startTimer"
Dim db As New cDebugReporter
    db.Report caller:=myName

    If mWinTimer Is Nothing Then Set mWinTimer = New cWinTimer
    mWinTimer.Startit pulseTime
    db.ExitMessage = "PulseTime = " & pulseTime
End Sub
Public Sub stopTimer()
Const myName As String = "cfloatHeaders.stopTimer"
Dim retVal As Long
Dim db As New cDebugReporter
    db.Report caller:=myName

    retVal = mWinTimer.Stopit
    db.ExitMessage = retVal
End Sub
Private Sub Class_Initialize()
Const myName As String = "cfloatHeaders.Class_Initialize"
Dim db As New cDebugReporter
    db.Report caller:=myName
        
    If mActivatedSheets Is Nothing Then
        Set mActivatedSheets = New Collection
    End If
    
    If mHostSheet Is Nothing Then Set mHostSheet = New cActiveSheet
    
    With mHostSheet
        .QualifyMethod = DEFisActiveSheet
        db.Report Message:=.ws.name
        If .isQualified Then
            db.Report Message:=.ws.name & " is active sheet"
            Call rangePic(DEFcolHdrRange, CHkey)
            mRowsPicWidth = rangePic(DEFrowHdrRange, RHkey)
            addActiveSheet .ws
            alignWithWindow True
            startTimer pulseTime
        End If
    End With
    
End Sub

Private Sub Class_Terminate()
Const myName As String = "cfloatHeaders.Class_Terminate"
Dim ws As Worksheet
Dim db As New cDebugReporter
    db.Report caller:=myName

    On Error Resume Next
    
    For Each ws In mActivatedSheets
        clearPics ws
        mActivatedSheets.Remove ws.name
    Next ws
    Set mWinTimer = Nothing
    Set mHostSheet = Nothing

End Sub

Private Sub mHostSheet_gotFocus(ByVal Wn As Window)
Const myName As String = "cfloatHeaders.mHostSheet_gotFocus"
Dim db As New cDebugReporter
    db.Report caller:=myName

    Class_Initialize
    
    db.ExitMessage = IIf(mWS Is Nothing, "NOT SET", "TIMER ACTIVATED")
End Sub

Private Sub mHostSheet_lostFocus(Cancel As Boolean)
Const myName As String = "cfloatHeaders.mHostSheet_lostFocus"
Dim db As New cDebugReporter
    db.Report caller:=myName


    Set mWinTimer = Nothing
    mHostSheet.ws = Nothing
    
End Sub
Private Sub mWinTimer_OnTimer()
Const myName As String = "cfloatHeaders.mWinTimer_OnTimer"
Dim db As New cDebugReporter
    db.Report caller:=myName

    With mHostSheet
        If .isQualified Then
            On Error Resume Next
            alignWithWindow
            On Error GoTo 0
        Else
            stopTimer
        End If
    End With
    
End Sub

