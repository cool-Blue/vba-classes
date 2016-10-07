VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'/////////////////////////////////////////////////////////////////////////////////////////////////////
'/Author cool.blue@y7mail.com
'/Nov 2014
'/Uses a Style statement stored in the shapes AlernativeText property.
'/If AlternativeText does not conform to the correct syntax,
'/  Writes a style statement registering size and position
'/  Syntax: <style>position: Top Left Width Height</style>
'/Else
'/  Reads the style and resets the shape size and position
'/////////////////////////////////////////////////////////////////////////////////////////////////////

Private mChart As Chart
Private mSheet As Worksheet

Public Property Let Parent(RHS As Object)
Dim s As Shape, os As cShape

  Select Case TypeName(RHS)
  Case "Chart"
    Set mChart = RHS
  Case "Worksheet"
    Set mSheet = RHS
  End Select
  
  For Each s In RHS.Shapes
    Set os = New cShape
    os.Shape = s
  Next s
  
End Property
