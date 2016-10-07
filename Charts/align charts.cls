VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub alignCharts()
Dim oCht As ChartObject, oOverCht As ChartObject
Dim Lbase As Long, Rovr As Long, W As Long, Wca As Long


  For Each oCht In Me.ChartObjects
    If oCht.Name = "BaseChart" Then Exit For
  Next oCht
  
  For Each oOverCht In Me.ChartObjects
    If oOverCht.Name = "OverlayChart" Then Exit For
  Next oOverCht
  Lbase = oCht.Chart.PlotArea.Left
  Wca = oCht.Chart.ChartArea.Width
  Rovr = Wca - oOverCht.Chart.PlotArea.Left - oOverCht.Chart.PlotArea.Width
  W = Wca - Lbase - Rovr
    
    
  With oOverCht
    .Top = oCht.Top
    .Left = oCht.Left
    .Width = oCht.Width
    .Height = oCht.Height
  End With
  
  With oCht.Chart
    .PlotArea.Width = W
  End With
  
  With oOverCht
    With .Chart.ChartArea
      .Top = oCht.Chart.ChartArea.Top
      .Left = oCht.Chart.ChartArea.Left
      .Height = oCht.Chart.ChartArea.Height
      '.Width = oCht.Chart.ChartArea.Width
    End With
    With .Chart.PlotArea
      .Top = oCht.Chart.PlotArea.Top
      .Left = oCht.Chart.PlotArea.Left + oCht.Chart.Axes(xlValue, xlPrimary).Width
      .Height = oCht.Chart.PlotArea.Height
      .Width = oCht.Chart.PlotArea.Width - oCht.Chart.Axes(xlValue, xlPrimary).Width
    End With
  End With
  oCht.Chart.PlotArea.Width = oCht.Chart.PlotArea.Width - oOverCht.Chart.Axes(xlValue, xlSecondary).Width
End Sub

Sub alignChartsOld()
Dim oCht As ChartObject, oOverCht As ChartObject
Dim Lbase As Long, Rovr As Long, W As Long, Wca As Long


  For Each oCht In Me.ChartObjects
    If oCht.Name = "BaseChart" Then Exit For
  Next oCht
  
  For Each oOverCht In Me.ChartObjects
    If oOverCht.Name = "OverlayChart" Then Exit For
  Next oOverCht
  Lbase = oCht.Chart.PlotArea.Left
  Wca = oCht.Chart.ChartArea.Width
  Rovr = Wca - oOverCht.Chart.PlotArea.Left - oOverCht.Chart.PlotArea.Width
  'W = Rovr - Lbase
  W = Wca - Lbase - Rovr
    
    
  With oOverCht
    .Top = oCht.Top
    .Left = oCht.Left
    .Width = oCht.Width
    .Height = oCht.Height
  End With
  
  With oCht.Chart
    .PlotArea.Width = .ChartArea.Width - .PlotArea.Left
    .PlotArea.Width = .PlotArea.Width - oOverCht.Chart.Axes(xlValue, xlSecondary).Width
  End With
  
  With oOverCht
    With .Chart.ChartArea
      .Top = oCht.Chart.ChartArea.Top
      .Left = oCht.Chart.ChartArea.Left
      .Height = oCht.Chart.ChartArea.Height
      .Width = oCht.Chart.ChartArea.Width
    End With
    With .Chart.PlotArea
      .Top = oCht.Chart.PlotArea.Top
      .Left = oCht.Chart.PlotArea.Left + oCht.Chart.Axes(xlValue, xlPrimary).Width
      .Height = oCht.Chart.PlotArea.Height
      .Width = oCht.Chart.PlotArea.Width - oCht.Chart.Axes(xlValue, xlPrimary).Width
    End With
  End With
  oCht.Chart.PlotArea.Width = oCht.Chart.PlotArea.Width - oOverCht.Chart.Axes(xlValue, xlSecondary).Width
  
End Sub

Private Sub Worksheet_Activate()
alignCharts

End Sub

