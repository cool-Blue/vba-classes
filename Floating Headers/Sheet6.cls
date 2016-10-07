VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/ © cool.blue@y7mail.com
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Option Explicit

'In Worksheet Module
Const cPulseTime As Long = 200

Property Get floatingHeaders() As Boolean
    floatingHeaders = True
End Property
Property Get isActiveSheet() As Boolean
    isActiveSheet = False
End Property
Property Get pulseTime() As Long
    pulseTime = cPulseTime
End Property

