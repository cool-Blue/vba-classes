VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FILTERS As String = "Text Files (*.txt)" & vbNull & "*.txt" & vbNull & _
                                  "All Files (*.*)" & vbNull & "*.*" & vbNull
Private Const FLAGS As Long = 0
Private Const DEF_FILENAME = "Log"

Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
' You won't use these.
'private Const OFN_ENABLEHOOK = &H20
'private Const OFN_ENABLETEMPLATE = &H40
'private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000
' New for Windows 95
Private Const OFN_EXPLORER = &H80000
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000

Dim mfso As FileSystemObject, mf As TextStream
Dim mSavePath As String

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
                  (pOpenfilename As gFILE) As Boolean
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" _
                  (pOpenfilename As gFILE) As Boolean

Private Type gFILE
 lStructSize As Long
 hwndOwner As Long
 hInstance As Long
 lpstrFilter As String
 lpstrCustomFilter As String
 nMaxCustFilter As Long
 nFilterIndex As Long
 lpstrFile As String
 nMaxFile As Long
 lpstrFileTitle As String
 nMaxFileTitle As Long
 lpstrInitialDir As String
 lpstrTitle As String
 FLAGS As Long
 nFileOffset As Integer
 nFileExtension As Integer
 lpstrDefExt As String
 lCustData As Long
 lpfnHook As Long
 lpTemplateName As String
End Type

Private Function FileToSave(Optional StartLookIn = vbNullString) As String
Dim OFN As gFILE
Dim path As String
Dim fileName As String
Dim a As Boolean
 
With OFN
  .lStructSize = Len(OFN)
  .lpstrFilter = FILTERS
  .lpstrFile = DEF_FILENAME & Space$(254 - Len(DEF_FILENAME))
  .nMaxFile = 255
  .lpstrFileTitle = Space$(254)
  .nMaxFileTitle = 255
  
  'Set Opening Folder stuff
  'If Not IsMissing(StartLookIn) Then
  ' .lpstrInitialDir = StartLookIn
  ''Else
  '' .lpstrInitialDir = "C:\CorelDrw Files"
  'End If
  
  .lpstrTitle = "Text File to dump trace"
  .FLAGS = OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
  .lpstrDefExt = "txt"
End With

a = GetSaveFileName(OFN)
If (a) Then
 path = Trim(OFN.lpstrFile)
 fileName = Trim(OFN.lpstrFileTitle)
Else
 path = ""
 fileName = ""
End If

FileToSave = path & "\" & fileName

End Function
Public Property Let NextLine(RHS As String)
  mf.WriteLine (RHS)
End Property
Private Sub Class_Initialize()
  mSavePath = FileToSave()
  Set mfso = New FileSystemObject
  Set mf = mfso.CreateTextFile(mSavePath)
End Sub

Private Sub Class_Terminate()
  mf.Close
End Sub
