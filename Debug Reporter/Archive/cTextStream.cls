VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTextStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FILTERS_DEF As String = "Text Files (*.txt)" & vbNull & "*.txt" & vbNull & _
                                  "All Files (*.*)" & vbNull & "*.*" & vbNull
Private Const FLAGS As Long = 0
Private Const FILENAME_DEF = "Log"

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

Private mfso As FileSystemObject, mf As TextStream
Private mSavePath As String
Private mFilters As String

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
Private Type ofnFilter
  ofnName As String
  ofnExt As String
End Type

Private saveDenied As Boolean

Private Function FileToSave(Optional StartLookIn = vbNullString) As String
Dim OFN As gFILE
Dim path As String
Dim fileName As String
Dim a As Boolean
 
With OFN
  .lStructSize = Len(OFN)
  .lpstrFilter = FILTERS_DEF
  .lpstrFile = FILENAME_DEF & Space$(254 - Len(FILENAME_DEF))
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
  FileToSave = path & "\" & fileName
Else
  FileToSave = vbNullString
End If


End Function
Public Sub addFilter(fExt As String, Optional fDesc As String = vbNullString, _
                      Optional insert As Boolean = False)
'Entries are not validated
Dim f As ofnFilter, vf As Variant
  If insert Then
    mFilters = fDesc & Chr(0) & fExt & Chr(0) & mFilters
  Else
    mFilters = mFilters & Chr(0) & fDesc & Chr(0) & fExt
  End If
  Debug.Print f.ofnName
End Sub
Public Property Let NextLine(RHS As String)
  If mf Is Nothing And Not saveDenied Then Create
  If Not mf Is Nothing Then mf.WriteLine RHS
End Property
Public Sub Create()
  mSavePath = FileToSave()
  If mSavePath <> vbNullString Then
    Set mfso = New FileSystemObject
    Set mf = mfso.CreateTextFile(mSavePath)
  Else
    saveDenied = True
  End If
End Sub
Private Sub Class_Initialize()
  mFilters = FILTERS_DEF
  Create
End Sub

Private Sub Class_Terminate()
  mf.Close
End Sub
