VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Series Formula and Inex are now globals and only read once when Index is specified
'Indx property added and series formula routines modified to accomodate this
'This Class module provides an easy way to access the items in a
'chart's SERIES formula. It can be exported and then imported into
'any project

'Developed by John Walkenbach, JWALK AND ASSOCIATES
'Copyright 1999. All rights reserved.
'May be used and distributed freely, but may not be sold.
'http://www.j-walk.com/ss/

'===================================
' PROPERTIES FOR CHARTSERIES OBJECT
'===================================

'Chart (read/write)
'myIndx (read/write)

'SeriesName (read/write)
'SeriesNameinFormula (write)
'XValues (read/write)
'Values (read/write)
'ValuesInFormula (write)
'PlotOrder (read/write)

'SeriesNameType (read-only)
'XValuesType (read-only)
'ValuesType (read-only)
'PlotOrderType (read-only)


'Sub ExampleUsage()
'    Dim MySeries As New ChartSeries
'    With MySeries
'        .Chart = Sheets(1).ChartObjects(1).Chart
'        .myIndx = 1
'        If .XValuesType = "Range" Then
'            MsgBox .XValues.Address
'        Else
'            MsgBox .XValues
'        End If
'    End With
'End Sub

Option Explicit
Const actRetrieve = 0
Const actReplace = 1
'Public mChartEventsClassParent As cChartEvents
Private mCurrChart As Chart    'accessible to all procedures
Private mSeriesCollection As Collection
Private mSeriesPoints As Collection
Dim currSeries As series
Dim mIndx As Integer 'accessible to all procedures
Dim SeriesFormula As String

Public LastOPerationSuccessful As Boolean

Property Let Indx(i As Integer)
    mIndx = i
'   Get the SERIES formula
    Set currSeries = mCurrChart.SeriesCollection(mIndx)
    SeriesFormula = currSeries.Formula
End Property
Property Get Indx() As Integer
    Indx = mIndx
End Property
Property Get Chart() As Chart
    Set Chart = mCurrChart
End Property

Property Let Chart(cht)
Dim s As Long
  Set mCurrChart = cht
  Set mSeriesCollection = New Collection
  For s = 1 To mCurrChart.SeriesCollection.Count
    Indx = s
    mSeriesCollection.Add currSeries, SeriesName
  Next s
  Indx = 1
End Property
Property Get series(Optional s As Variant) As series
  If Not IsMissing(s) Then CurrentSeries = s
  Set series = currSeries
End Property
Property Get SeriesName() As Variant
    If SeriesNameType = "Range" Then
        Set SeriesName = Range(SERIESFormulaElement(mCurrChart, 1))
    Else
        SeriesName = SERIESFormulaElement(mCurrChart, 1)
    End If
End Property
Property Let CurrentSeries(RHS As Variant)

  Select Case TypeName(RHS)
  Case "String"
    Indx = seriesIndexByName(CStr(RHS))
  Case "Long", "Integer"
    Indx = RHS
  Case Else
  End Select
  
End Property
Property Get Item(s As Variant) As series
Attribute Item.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Item.VB_UserMemId = 0
  
  Select Case TypeName(s)
  Case "String"
    Set Item = mSeriesCollection(CStr(s))
  Case "Long", "Integer"
    Set Item = mSeriesCollection(s)
  Case Else
  End Select
  
End Property
Property Let SeriesNameinFormula(sName As String)
    LastOPerationSuccessful = setSERIESFormulaElement(mCurrChart, 1, sName)
End Property
Property Let SeriesName(sName)
    mCurrChart.SeriesCollection(mIndx).name = sName
End Property
Property Get SeriesNameType() As String
    SeriesNameType = SERIESFormulaElementType(mCurrChart, 1)
End Property
Property Get SeriesPoints()
    Set SeriesPoints = currSeries.Points
End Property
Property Get dataLabelTxt(i As Long) As String
    dataLabelTxt = currSeries.DataLabels(i).caption
End Property
Property Get dataLabel(i As Long) As dataLabel
    Set dataLabel = currSeries.DataLabels(i)
End Property
Property Get XValues() As Variant
    If XValuesType = "Range" Then
        Set XValues = Range(SERIESFormulaElement(mCurrChart, 2))
    Else
        XValues = SERIESFormulaElement(mCurrChart, 2)
    End If
End Property
Property Let XValues(XVals)
    mCurrChart.SeriesCollection(mIndx).XValues = XVals
End Property
Property Get XValuesType() As String
    XValuesType = SERIESFormulaElementType(mCurrChart, 2)
End Property
Property Get Values() As Variant
    If ValuesType = "Range" Then
        Set Values = Range(SERIESFormulaElement(mCurrChart, 3))
    Else
        Values = SERIESFormulaElement(mCurrChart, 3)
    End If
End Property
Property Let Values(Vals)
    mCurrChart.SeriesCollection(mIndx).Values = Vals
End Property
Property Let ValuesInFormula(Vals As String)
    LastOPerationSuccessful = setSERIESFormulaElement(mCurrChart, 3, Vals)
End Property
Property Get ValuesType() As String
    ValuesType = SERIESFormulaElementType(mCurrChart, 3)
End Property
Property Get PlotOrder()
        PlotOrder = SERIESFormulaElement(mCurrChart, 4)
End Property
Property Let PlotOrder(PltOrder)
    mCurrChart.SeriesCollection(mIndx).PlotOrder = PltOrder
End Property
Property Get PlotOrderType() As String
    PlotOrderType = SERIESFormulaElementType(mCurrChart, 4)
End Property
Property Let sType(SeriesType As Long)
Dim s As series
    Set s = mCurrChart.SeriesCollection(mIndx)
    With s
        .Type = SeriesType
        Select Case SeriesType
        Case xlArea
        Case xlLine
            .Border.LineStyle = xlNone 'xlHairline 'xlContinuous 'xlNone
        Case xlColumn
        Case Else
        End Select
    End With
End Property
Property Get sType() As Long
    sType = currSeries.Type
End Property
Property Get chartType() As Long
  chartType = currSeries.chartType
End Property
Property Get sTypeText() As String
    Select Case sType
    Case xlArea
        sTypeText = "Area"
    Case xlLine
        sTypeText = "Line"
    Case xlColumn
        sTypeText = "Column"
    Case Else
        sTypeText = "Other"
    End Select
End Property
Property Let AxisGroup(ax As Long)
With mCurrChart.SeriesCollection(mIndx)
    Select Case ax
    Case 1
        .AxisGroup = xlPrimary
    Case 2
        .AxisGroup = xlSecondary
    End Select
End With
End Property
Property Get catAxis(ag As Long) As Axis
With mCurrChart
    Set catAxis = .Axes(xlCategory, ag)
End With
End Property
Property Let sTypeText(t As String)
    Select Case t
    Case "Area"
        sType = xlArea
    Case "Line"
        sType = xlLine
    Case "Column"
        sType = xlColumn
    Case Else
    End Select
End Property
Property Let LineColour(Colour)
With mCurrChart.SeriesCollection(mIndx)
    On Error Resume Next
    .Format.Line.Visible = msoTrue
    .Format.Line.ForeColor.RGB = Colour
    '.Border.LineStyle = xlNone 'xlHairline 'xlContinuous 'xlNone
    '.Border.Color = Colour
    '.Interior.Color = Colour
    On Error GoTo 0
End With
End Property
Property Get LineColour()
With mCurrChart.SeriesCollection(mIndx)
    LineColour = .Format.Line.ForeColor
End With
End Property
Property Let FillColour(Colour)
With mCurrChart.SeriesCollection(mIndx)
    Select Case .Type
    Case xlArea
        .Format.Fill.Visible = msoTrue
        .Format.Fill.ForeColor.RGB = Colour
    Case xlLine
        LineColour = Colour
    Case Else
    End Select
End With
End Property
Property Get FillColour()
With currSeries
    Select Case .Type
    Case xlArea
        FillColour = .Format.Fill.ForeColor
    Case xlLine
        FillColour = .Format.Line.ForeColor
    Case Else
    End Select
End With
End Property
Property Get seriesPattern() As Long
With currSeries.Format.Fill
    Select Case mCurrChart.SeriesCollection(mIndx).Type
    Case xlArea
        seriesPattern = .pattern
    Case xlLine
        seriesPattern = .pattern
    End Select
End With
End Property
Property Let seriesPattern(p As Long)
With mCurrChart.SeriesCollection(mIndx).Format.Fill
'    .Visible = True
    If p <> 0 Then
        Select Case mCurrChart.SeriesCollection(mIndx).Type
        Case xlArea
            .Patterned p
        Case xlLine
        End Select
    Else
        .Solid
    End If
End With
End Property
Property Let Transparency(t)
With mCurrChart.SeriesCollection(mIndx)
    Select Case .Type
    Case xlArea
        .Format.Fill.Transparency = t
    Case xlLine
        .Format.Line.Transparency = t
    Case Else
    End Select
End With
End Property
Property Get Transparency()
With currSeries
    Select Case .Type
    Case xlArea
        Transparency = .Format.Fill.Transparency
    Case xlLine
        Transparency = .Format.Line.Transparency
    Case Else
    End Select
End With
End Property
Property Let LineWeight(w)
With mCurrChart.SeriesCollection(mIndx)
    .Format.Line.Weight = w
End With
End Property
Property Get LineWeight()
With currSeries
    LineWeight = .Format.Line.Weight
End With
End Property
Property Let LineVisible(v)
With mCurrChart.SeriesCollection(mIndx)
    .Format.Line.Visible = v
End With
End Property
Public Sub noMarkers()
    With mCurrChart.SeriesCollection(mIndx)
        .MarkerStyle = xlMarkerStyleNone
'        .Visible = msoTrue
    End With
End Sub
Public Function sTypeFromText(t As String) As Long
    Select Case t
    Case "Area"
        sTypeFromText = xlArea
    Case "Line"
        sTypeFromText = xlLine
    Case "Column"
        sTypeFromText = xlColumn
    Case Else
    End Select
End Function

Private Function SERIESFormulaElementType(ChartObj, Element) As String
'   Returns a string that describes the element of a chart's SERIES formula
'   This function essentially parses and analyzes a SERIES formula

'   Element 1: Series Name. Returns "Range" , "Empty", or "String"
'   Element 2: XValues. Returns "Range", "Empty", or "Array"
'   Element 3: Values. Returns "Range" or "Array"
'   Element 4: PlotOrder. Always returns "Integer"

'   Globals:
'   Dim mIndx As Integer 'accessible to all procedures
'   Dim SeriesFormula As String

    Dim FirstComma As Integer, SecondComma As Integer, LastComma As Integer
    Dim FirstParen As Integer, SecondParen As Integer
    Dim FirstBracket As Integer, SecondBracket As Integer
    Dim StartY As Integer
    Dim SeriesName, XValues, Values, PlotOrder As Integer
    
'   Exit if Surface chart (surface chrarts do not have SERIES formulas)
    If ChartObj.chartType >= 83 And ChartObj.chartType <= 86 Then
        SERIESFormulaElementType = "ERROR - SURFACE CHART"
        Exit Function
    End If
    
'   Exit if nonexistent series is specified
    If mIndx > ChartObj.SeriesCollection.Count Or mIndx < 1 Then
        SERIESFormulaElementType = "ERROR - BAD SERIES NUMBER"
        Exit Function
    End If

'   Exit if element is > 4
    If Element > 4 Or Element < 1 Then
        SERIESFormulaElementType = "ERROR - BAD ELEMENT NUMBER"
        Exit Function
    End If


'   Get the First Element (Series Name)
    FirstParen = InStr(1, SeriesFormula, "(")
    FirstComma = InStr(1, SeriesFormula, ",")
    SeriesName = Mid(SeriesFormula, FirstParen + 1, FirstComma - FirstParen - 1)
    If Element = 1 Then
        If IsRange(SeriesName) Then
            SERIESFormulaElementType = "Range"
        Else
            If SeriesName = "" Then
                SERIESFormulaElementType = "Empty"
            Else
                If TypeName(SeriesName) = "String" Then
                    SERIESFormulaElementType = "String"
                End If
            End If
        End If
        Exit Function
    End If

'   Get the Second Element (X Range)
    If Mid(SeriesFormula, FirstComma + 1, 1) = "(" Then
'       Multiple ranges
        FirstParen = FirstComma + 2
        SecondParen = InStr(FirstParen, SeriesFormula, ")")
        XValues = Mid(SeriesFormula, FirstParen, SecondParen - FirstParen)
        StartY = SecondParen + 1
    Else
        If Mid(SeriesFormula, FirstComma + 1, 1) = "{" Then
'           Literal Array
            FirstBracket = FirstComma + 1
            SecondBracket = InStr(FirstBracket, SeriesFormula, "}")
            XValues = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            StartY = SecondBracket + 1
        Else
'          A single range
            SecondComma = InStr(FirstComma + 1, SeriesFormula, ",")
            XValues = Mid(SeriesFormula, FirstComma + 1, SecondComma - FirstComma - 1)
            StartY = SecondComma
        End If
    End If
    If Element = 2 Then
        If IsRange(XValues) Then
            SERIESFormulaElementType = "Range"
        Else
            If XValues = "" Then
                SERIESFormulaElementType = "Empty"
            Else
                SERIESFormulaElementType = "Array"
            End If
        End If
        Exit Function
    End If

'   Get the Third Element (Y Range)
    If Mid(SeriesFormula, StartY + 1, 1) = "(" Then
'       Multiple ranges
        FirstParen = StartY + 1
        SecondParen = InStr(FirstParen, SeriesFormula, ")")
        Values = Mid(SeriesFormula, FirstParen + 1, SecondParen - FirstParen - 1)
        LastComma = SecondParen + 1
    Else
        If Mid(SeriesFormula, StartY + 1, 1) = "{" Then
'           Literal Array
            FirstBracket = StartY + 1
            SecondBracket = InStr(FirstBracket, SeriesFormula, "}")
            Values = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            LastComma = SecondBracket + 1
        Else
'          A single range
            FirstComma = StartY
            SecondComma = InStr(FirstComma + 1, SeriesFormula, ",")
            Values = Mid(SeriesFormula, FirstComma + 1, SecondComma - FirstComma - 1)
            LastComma = SecondComma
        End If
    End If
    If Element = 3 Then
        If IsRange(Values) Then
            SERIESFormulaElementType = "Range"
        Else
            SERIESFormulaElementType = "Array"
        End If
        Exit Function
    End If
    
'   Get the Fourth Element (Plot Order)
    PlotOrder = Mid(SeriesFormula, LastComma + 1, Len(SeriesFormula) - LastComma - 1)
    If Element = 4 Then
        SERIESFormulaElementType = "Integer"
        Exit Function
    End If
End Function
Private Function SERIESFormulaElement(ChartObj, Element) As String
'   Returns one of four elements in a chart's SERIES formula (as a string)
'   This function essentially parses and analyzes a SERIES formula

'   Element 1: Series Name. Can be a range reference, a literal value, or nothing
'   Element 2: XValues. Can be a range reference (including a non-contiguous range), a literal array, or nothing
'   Element 3: Values. Can be a range reference (including a non-contiguous range), or a literal array
'   Element 4: PlotOrder. Must be an integer

'   Globals:
'   Dim mIndx As Integer 'accessible to all procedures
'   Dim SeriesFormula As String
    
    Dim FirstComma As Integer, SecondComma As Integer, LastComma As Integer
    Dim FirstParen As Integer, SecondParen As Integer
    Dim FirstBracket As Integer, SecondBracket As Integer
    Dim StartY As Integer
    Dim SeriesName, XValues, Values, PlotOrder As Integer, NewFormula As String
    
'   Exit if Surface chart (surface chrarts do not have SERIES formulas)
    If ChartObj.chartType >= 83 And ChartObj.chartType <= 86 Then
        SERIESFormulaElement = "ERROR - SURFACE CHART"
        Exit Function
    End If
    
'   Exit if nonexistent series is specified
    If mIndx > ChartObj.SeriesCollection.Count Or mIndx < 1 Then
        SERIESFormulaElement = "ERROR - BAD SERIES NUMBER"
        Exit Function
    End If

'   Exit if element is > 4
    If Element > 4 Then
        SERIESFormulaElement = "ERROR - BAD ELEMENT NUMBER"
        Exit Function
    End If


'   Get the First Element (Series Name)
    FirstParen = InStr(1, SeriesFormula, "(")
    FirstComma = InStr(1, SeriesFormula, ",")
    SeriesName = Mid(SeriesFormula, FirstParen + 1, FirstComma - FirstParen - 1)
    If Element = 1 Then
        SERIESFormulaElement = SeriesName
        Exit Function
    End If

'   Get the Second Element (X Range)
    If Mid(SeriesFormula, FirstComma + 1, 1) = "(" Then
'       Multiple ranges
        FirstParen = FirstComma + 2
        SecondParen = InStr(FirstParen, SeriesFormula, ")")
        XValues = Mid(SeriesFormula, FirstParen, SecondParen - FirstParen)
        StartY = SecondParen + 1
    Else
        If Mid(SeriesFormula, FirstComma + 1, 1) = "{" Then
'           Literal Array
            FirstBracket = FirstComma + 1
            SecondBracket = InStr(FirstBracket, SeriesFormula, "}")
            XValues = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            StartY = SecondBracket + 1
        Else
'          A single range
            SecondComma = InStr(FirstComma + 1, SeriesFormula, ",")
            XValues = Mid(SeriesFormula, FirstComma + 1, SecondComma - FirstComma - 1)
            StartY = SecondComma
        End If
    End If
    If Element = 2 Then
        SERIESFormulaElement = XValues
        Exit Function
    End If

'   Get the Third Element (Y Range)
    If Mid(SeriesFormula, StartY + 1, 1) = "(" Then
'       Multiple ranges
        FirstParen = StartY + 1
        SecondParen = InStr(FirstParen, SeriesFormula, ")")
        Values = Mid(SeriesFormula, FirstParen + 1, SecondParen - FirstParen - 1)
        LastComma = SecondParen + 1
    Else
        If Mid(SeriesFormula, StartY + 1, 1) = "{" Then
'           Literal Array
            FirstBracket = StartY + 1
            SecondBracket = InStr(FirstBracket, SeriesFormula, "}")
            Values = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            LastComma = SecondBracket + 1
        Else
'          A single range
            FirstComma = StartY
            SecondComma = InStr(FirstComma + 1, SeriesFormula, ",")
            Values = Mid(SeriesFormula, FirstComma + 1, SecondComma - FirstComma - 1)
            LastComma = SecondComma
        End If
    End If
    If Element = 3 Then
        SERIESFormulaElement = Values
        Exit Function
    End If
    
'   Get the Fourth Element (Plot Order)
    PlotOrder = Mid(SeriesFormula, LastComma + 1, Len(SeriesFormula) - LastComma - 1)
    If Element = 4 Then
        SERIESFormulaElement = PlotOrder
        Exit Function
    End If
End Function
Private Function setSERIESFormulaElement(ChartObj, Element, NewElement As String) As Boolean
'   Returns one of four elements in a chart's SERIES formula (as a string)
'   This function essentially parses and analyzes a SERIES formula

'   Element 1: Series Name. Can be a range reference, a literal value, or nothing
'   Element 2: XValues. Can be a range reference (including a non-contiguous range), a literal array, or nothing
'   Element 3: Values. Can be a range reference (including a non-contiguous range), or a literal array
'   Element 4: PlotOrder. Must be an integer

'   Globals:
'   Dim mIndx As Integer 'accessible to all procedures
'   Dim SeriesFormula As String

    Dim FirstComma As Integer, SecondComma As Integer, LastComma As Integer
    Dim FirstParen As Integer, SecondParen As Integer
    Dim FirstBracket As Integer, SecondBracket As Integer
    Dim StartY As Integer
    Dim SeriesName, XValues, Values, PlotOrder As Integer, NewFormula As String
    
'   Exit if Surface chart (surface chrarts do not have SERIES formulas)
    If ChartObj.chartType >= 83 And ChartObj.chartType <= 86 Then
        setSERIESFormulaElement = False
        Exit Function
    End If
    
'   Exit if nonexistent series is specified
    If mIndx > ChartObj.SeriesCollection.Count Or mIndx < 1 Then
        setSERIESFormulaElement = False
        Exit Function
    End If

'   Exit if element is > 4
    If Element > 4 Then
        setSERIESFormulaElement = False
        Exit Function
    End If

'   Get the First Element (Series Name)
    FirstParen = InStr(1, SeriesFormula, "(")
    FirstComma = InStr(1, SeriesFormula, ",")
    SeriesName = Mid(SeriesFormula, FirstParen + 1, FirstComma - FirstParen - 1)
    If Element = 1 Then
        SeriesFormula = Left(SeriesFormula, FirstParen) & NewElement & Right(SeriesFormula, Len(SeriesFormula) - FirstComma + 1)
        ChartObj.SeriesCollection(mIndx).Formula = SeriesFormula
        setSERIESFormulaElement = True
        Exit Function
    End If

'   Get the Second Element (X Range)
    If Mid(SeriesFormula, FirstComma + 1, 1) = "(" Then
'       Multiple ranges
        FirstParen = FirstComma + 2
        SecondParen = InStr(FirstParen, SeriesFormula, ")")
        StartY = SecondParen + 1
    Else
        If Mid(SeriesFormula, FirstComma + 1, 1) = "{" Then
'           Literal Array
            FirstBracket = FirstComma + 1
            SecondBracket = InStr(FirstBracket, SeriesFormula, "}")
            StartY = SecondBracket + 1
        Else
'          A single range
            SecondComma = InStr(FirstComma + 1, SeriesFormula, ",")
            StartY = SecondComma
        End If
    End If
    If Element = 2 Then
        SeriesFormula = Left(SeriesFormula, FirstComma) & NewElement & Right(SeriesFormula, Len(SeriesFormula) - StartY + 1)
        ChartObj.SeriesCollection(mIndx).Formula = SeriesFormula
        setSERIESFormulaElement = True
        Exit Function
    End If

'   Get the Third Element (Y Range)
    If Mid(SeriesFormula, StartY + 1, 1) = "(" Then
'       Multiple ranges
        FirstParen = StartY + 1
        SecondParen = InStr(FirstParen, SeriesFormula, ")")
        Values = Mid(SeriesFormula, FirstParen + 1, SecondParen - FirstParen - 1)
        LastComma = SecondParen + 1
    Else
        If Mid(SeriesFormula, StartY + 1, 1) = "{" Then
'           Literal Array
            FirstBracket = StartY + 1
            SecondBracket = InStr(FirstBracket, SeriesFormula, "}")
            Values = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            LastComma = SecondBracket + 1
        Else
'          A single range
            FirstComma = StartY
            SecondComma = InStr(FirstComma + 1, SeriesFormula, ",")
            LastComma = SecondComma
        End If
    End If
    If Element = 3 Then
        SeriesFormula = Left(SeriesFormula, StartY) & NewElement & Right(SeriesFormula, Len(SeriesFormula) - SecondComma + 1)
        ChartObj.SeriesCollection(mIndx).Formula = SeriesFormula
        setSERIESFormulaElement = True
        Exit Function
    End If
    
'   Get the Fourth Element (Plot Order)
    PlotOrder = Mid(SeriesFormula, LastComma + 1, Len(SeriesFormula) - LastComma - 1)
    If Element = 4 Then
        SeriesFormula = Left(SeriesFormula, LastComma) & NewElement & ")"
        ChartObj.SeriesCollection(Indx).Formula = SeriesFormula
        setSERIESFormulaElement = True
        Exit Function
    End If
End Function
Private Function IsRange(ref) As Boolean
'   Returns True if ref is a Range
    Dim x As Range
    On Error Resume Next
    Set x = Range(ref)
    If Err = 0 Then
        IsRange = True
    Else
        IsRange = False
        Err.Clear
    End If
End Function
Private Function seriesIndexByName(sName As String) As Long
Dim s As series, os As ChartSeries, i As Long

  Set os = New ChartSeries
  os.Chart = Chart
  For i = 1 To os.Chart.SeriesCollection.Count
    With os
      .Indx = i
      If os.SeriesName = sName Then Exit For
    End With
  Next i
  seriesIndexByName = i
  
End Function
Private Sub Class_Terminate()
    Debug.Print mCurrChart.name & " Series Class Terminated"
    Set mCurrChart = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Enumerator: Edit in Notepad to see this.
'    Attribute NewEnum.VB_UserMemId = -4
'    Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mSeriesCollection.[_NewEnum]
End Property
