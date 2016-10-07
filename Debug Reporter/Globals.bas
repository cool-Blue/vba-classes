Attribute VB_Name = "Globals"
Option Explicit

Public Const gcDebugMode As Boolean = True
Public Const gdebugOutoutToFile As Boolean = True
Public gCallDepth As Long
Public glogFile As cTextStream
Public gSavState(0 To 255) As Byte

