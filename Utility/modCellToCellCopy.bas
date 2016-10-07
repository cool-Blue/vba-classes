Attribute VB_Name = "modCellToCellCopy"
Option Explicit
Const form = "#0.0E+0"
Const formx = "0.000000"
Sub testcopy()
Dim i As Long, t0 As Double
Dim dest As Worksheet, src As Worksheet, v As Variant
    Set dest = ThisWorkbook.Sheets("Sheet4")
    Set src = ThisWorkbook.Sheets("Sheet3")
    t0 = MicroTimer
    For i = 1 To 900
        dest.Cells(1, i).value = src.Cells(1, i).value
    Next i
    Debug.Print "cells:         " & vbTab & Application.WorksheetFunction.Text(MicroTimer - t0, formx)
    t0 = MicroTimer
    dest.Cells(1, 1) = src.Cells(1, 1).Resize(1, 900)
    Debug.Print "range:         " & vbTab & Application.WorksheetFunction.Text(MicroTimer - t0, formx)
    t0 = MicroTimer
    dest.Cells(1, 1).value = src.Cells(1, 1).Resize(1, 900).value
    Debug.Print "value:         " & vbTab & Application.WorksheetFunction.Text(MicroTimer - t0, formx)
    t0 = MicroTimer
    dest.Cells(1, 1).Value2 = src.Cells(1, 1).Resize(1, 900).Value2
    Debug.Print "value2:       " & vbTab & Application.WorksheetFunction.Text(MicroTimer - t0, formx)
    t0 = MicroTimer
    src.Cells(1, 1).Resize(1, 900).Copy (dest.Cells(1, 1))
    Debug.Print "copy dest:   " & vbTab & Application.WorksheetFunction.Text(MicroTimer - t0, formx)
    t0 = MicroTimer
    src.Cells(1, 1).Resize(1, 900).Copy
    dest.Paste (dest.Cells(1, 1))
    Debug.Print "copy paste:   " & vbTab & Application.WorksheetFunction.Text(MicroTimer - t0, formx)
    t0 = MicroTimer
    v = src.Cells(1, 1).Resize(1, 900).value
    dest.Cells(1, 1).value = v
    Debug.Print "variant:     " & vbTab & Application.WorksheetFunction.Text(MicroTimer - t0, formx)
    t0 = MicroTimer
    v = src.Cells(1, 1).Resize(1, 900).Value2
    dest.Cells(1, 1).Value2 = v
    Debug.Print "variant2:     " & vbTab & Application.WorksheetFunction.Text(MicroTimer - t0, formx)

End Sub
Sub test()
    Application.ScreenUpdating = False
    Debug.Print
    Debug.Print "Screenupdating False"
    testcopy
    Application.ScreenUpdating = True
    Debug.Print "Screenupdating True"
    testcopy
End Sub
