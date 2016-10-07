VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const switchOff As Boolean = False
Public Enum appOnOff
  appOn
  appOff
End Enum
Private WithEvents mApp As Application
Attribute mApp.VB_VarHelpID = -1

Private mScreenUpdateMode As appOnOff
Private mCalculationMode As XlCalculation

Private Sub Class_Initialize()
  Set mApp = Excel.Application
  
 mScreenUpdateMode = IIf(mApp.screenUpdating, appOn, appOff)
 mCalculationMode = mApp.calculation
  
End Sub
Public Property Let screenUpdateMode(RHS As appOnOff)
Const myName As String = "cApplication.screenUpdateMode"
Dim db As New cDebugReporter
Static pScreenUpdateMode As appOnOff

  Select Case RHS
  Case appOff
    pScreenUpdateMode = mApp.screenUpdating
    If pScreenUpdateMode Then mApp.screenUpdating = False
  Case appOn
    If pScreenUpdateMode Then mApp.screenUpdating = True
  End Select
  
  db.exitMessage = IIf(RHS = appOn, "Screen On", "Screen Off")
End Property

Public Property Let calculationMode(RHS As XlCalculation)
Const myName As String = "cApplication.calculationMode"
Dim db As New cDebugReporter
Dim exitMessage As String

  mApp.calculation = RHS
  
  Select Case RHS
  Case xlCalculationAutomatic
    exitMessage = "xlCalculationAutomatic"
  Case xlCalculationManual
    exitMessage = "xlCalculationManual"
  Case xlCalculationSemiautomatic
    exitMessage = "xlCalculationSemiautomatic"
  End Select
  db.exitMessage = exitMessage
End Property

Public Property Let eventMode(RHS As appOnOff)
Const myName As String = "cApplication.eventMode"
Dim db As New cDebugReporter

  Select Case RHS
  Case appOff
    mApp.enableEvents = False
  Case appOn
    mApp.enableEvents = True
  End Select
  
  db.exitMessage = IIf(RHS = appOn, "Events On", "Events Off")
End Property

Public Sub quietMode(Optional screenUpdating As appOnOff = appOff, _
                      Optional calcMode As XlCalculation = xlCalculationManual, _
                      Optional events As appOnOff = appOff)
                      
  If Not switchOff Then
    screenUpdateMode = screenUpdating
    calculationMode = calcMode
    eventMode = events
  End If
  
End Sub
Private Sub Class_Terminate()

  eventMode = appOn
  screenUpdateMode = mScreenUpdateMode
  calculationMode = mCalculationMode
  
End Sub
