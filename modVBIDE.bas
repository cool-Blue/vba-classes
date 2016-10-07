Attribute VB_Name = "modVBIDE"
Option Explicit

'Reference: Microsoft Visual Basic for Applications Extensibility 5.3
Sub removecomments()
Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent
Dim vbModLine As String, l As Long
    Set vbp = ThisWorkbook.VBProject
    With vbp
        For Each vbc In vbp.VBComponents
            With vbc
                Debug.Print .Name & vbTab & .Type & vbTab & "Lines: " & .CodeModule.CountOfLines & vbTab & "Prperties: " & .Properties.Count
                If .Type = vbext_ct_StdModule Then
                    With .CodeModule
                        For l = 1 To .CountOfLines
                            If l > .CountOfLines Then Exit For
                            vbModLine = .Lines(l, 1)
                            If (Left(Trim(vbModLine), 1) = "'") Or (Len(Trim(vbModLine)) = 0) Then
                                Debug.Print l & vbTab & "Deleted" & vbTab & vbModLine
                                .DeleteLines (l)
                                l = l - 1
                            End If
                        Next l
                    End With
                End If
            End With
        Next vbc
    End With
End Sub
Sub lineLengths()
Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent
Dim vbModLine As String, l As Long
    Set vbp = ThisWorkbook.VBProject
    With vbp
        For Each vbc In vbp.VBComponents
            With vbc
                If .Type = vbext_ct_StdModule Then
                    With .CodeModule
                        Debug.Print
                        For l = 1 To .CountOfLines
                            vbModLine = .Lines(l, 1)
                            Debug.Print l & vbTab & Len(Trim(vbModLine)) & vbTab & vbModLine
                        Next l
                    End With
                End If
            End With
        Next vbc
    End With
End Sub
Public Sub IndentComments(modName As String)
Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent
Dim vbModLine As String, l As Long
    Set vbp = ThisWorkbook.VBProject
    With vbp
        For Each vbc In vbp.VBComponents
            With vbc
                If .Type = vbext_ct_StdModule And .Name = modName Then
                    With .CodeModule
                        Debug.Print
                        For l = 1 To .CountOfLines
                            vbModLine = .Lines(l, 1)
                            If Left(vbModLine, 2) = "'/" Then
                                If Left(vbModLine, 3) <> "'//" Then
                                    vbModLine = "'/" & vbTab & Right(vbModLine, Len(vbModLine) - 2)
                                    Debug.Print vbModLine
                                    .DeleteLines l
                                    .InsertLines l, vbModLine
                                End If
                            End If
                        Next l
                    End With
                End If
            End With
        Next vbc
    End With
End Sub

