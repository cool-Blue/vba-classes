VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clSheetObjects_Variable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub init(NS As cldefSheetObjects_NameSpace, Name As String)

End Sub

Public Property Get NameSpaceObject() As cldefSheetObjects_NameSpace

End Property

Public Property Get Range() As Range

End Property

Public Property Get Name() As String

End Property

Public Property Get varType() As String

End Property

Public Property Get baseAddress() As String

End Property
Public Property Let baseAddress(RHS As String)

End Property

Public Property Get Contents() As String()

End Property
Public Property Let Contents(RHS() As String)

End Property

Public Property Get WordCount() As Long

End Property
'/////////////////////////////////////////////////////////////////////////////////////////////
