VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clStingsToColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'   Example Usage:
'   Sub formatCols()
'   Const colwidth As Long = 10
'   Dim Message As clStingsToColumns
'
'       Set Message = New clStingsToColumns
'       With Message
'           .tabs = Array(colwidth, colwidth, colwidth, colwidth)
'           .indents = 1
'           Debug.Print .Format("col1", "col2", "col3", "col4")
'       End With
'
'   End Sub
'   Sub formatCols1()
'   Const colwidth As Long = 10
'   Dim Message As clStingsToColumns
'
'       Set Message = New clStingsToColumns
'       With Message
'           .setCols colwidth, colwidth, colwidth, colwidth
'           .setIndents 1, 2, 1
'           Debug.Print .Format("col1", "col2", "col3", "col4")
'       End With
'
'   End Sub
Const DELIM As String = vbTab
Const defDiv As String = ":"
Private mWidths As Variant
Private mIndents As Variant
Private mwUb As Long
Private miUb As Long
Private mcUb As Long
Public Output As String
Public Divider As String
Property Let tabs(w As Variant)
    setVarray mwUb, mWidths, w
End Property
Property Let indents(i As Variant)
    setVarray miUb, mIndents, i
End Property
Public Sub setCols(ParamArray w())
Dim v As Variant
    v = w
    setVarray mwUb, mWidths, v
End Sub
Public Sub setIndents(ParamArray i())
Dim v As Variant
    v = i
    setVarray miUb, mIndents, v
End Sub
Private Sub setVarray(UB As Long, pElements As Variant, i As Variant)
    If IsArray(i) Then
        UB = UBound(i)
        ReDim pElements(UB)
        pElements = i
    Else
        UB = 0
        ReDim pElements(UB)
        pElements(0) = i
    End If
End Sub
Property Let Content(tabDelimContent As String)
Dim c() As String, s As String, i As Long, ind As Long, w As Long
Dim indPad As String, endPad As String

    c = Split(tabDelimContent, DELIM)
    mcUb = UBound(c)
    Output = ""
    indPad = ""
    For i = 0 To mcUb
        If i <= miUb Then indPad = String(mIndents(i), " ")
        If i <= mwUb Then w = mWidths(i)
        endPad = String(max(0, w - Len(c(i)) - Len(indPad) - Len(Divider)), " ")
        Output = Output & indPad & c(i) & endPad & Divider
    Next i
    
End Property
Public Function Format(ParamArray Strings()) As String
Dim i As Long, indPad As String, endPad As String, w As Long, retval As String

    mcUb = UBound(Strings)
    retval = ""
    indPad = ""
    For i = 0 To mcUb
        If i <= miUb Then indPad = String(mIndents(i), " ")
        If i <= mwUb Then w = mWidths(i)
        endPad = String(max(0, w - Len(Strings(i)) - Len(indPad) - Len(Divider)), " ")
        If IsNumeric(Strings(i)) Then
            retval = retval & endPad & Strings(i) & indPad & Divider
        Else
            retval = retval & indPad & Strings(i) & endPad & Divider
        End If
    Next i
    Format = retval
    
End Function

Private Sub Class_Initialize()
    Divider = defDiv
End Sub
