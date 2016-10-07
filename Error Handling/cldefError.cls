VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cldefError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mErr As ErrObject
Private mDecsription As String
Private mHelpContext As Long
Private mHelpFile As String
Private mLastDIIErroe As Long
Private mNumber As Long
Private mSource As String

Property Let Err(e As ErrObject)
    Set mErr = e
    With mErr
        mDecsription = .Description
        mHelpContext = .HelpContext
        mHelpFile = .HelpFile
        mLastDIIErroe = .LastDllError
        mNumber = .Number
        mSource = .Source
    End With 'mErr
End Property
Property Let Number(pNumber As Long)
    mNumber = pNumber
End Property
Property Get Number() As Long
Attribute Number.VB_UserMemId = 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Default value: Edit in Notepad to see this.
'Attribute Number.VB_UserMemId = 0
	Number = mNumber
End Property

Property Let Decsription(pDecsription As String): mDecsription = pDecsription: End Property
Property Let HelpContext(pHelpContext As Long): mHelpContext = pHelpContext: End Property
Property Let HelpFile(pHelpFile As String): mHelpFile = pHelpFile: End Property
Property Let LastDIIErroe(pLastDIIErroe As Long): mLastDIIErroe = pLastDIIErroe: End Property
Property Let Source(pSource As String): mSource = pSource: End Property

Property Get Decsription() As String: Decsription = mDecsription: End Property
Property Get HelpContext() As Long: HelpContext = mHelpContext: End Property
Property Get HelpFile() As String: HelpFile = mHelpFile: End Property
Property Get LastDIIErroe() As Long: LastDIIErroe = mLastDIIErroe: End Property
Property Get Source() As String: Source = mSource: End Property

