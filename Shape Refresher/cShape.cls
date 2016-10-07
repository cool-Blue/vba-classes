VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const styleMask As String = "^\<style\>\s*(.+\b)\s*\<\/style\>"
Const startTag As String = "<style>"
Const endTag As String = "</style>"
Const positionTag As String = "(position\s*:)\s+"
Const positionBody As String = "position\s*:\s*((\s+[\d\.]+){1,4})"

Enum positionType
  ptTop = 0
  ptLeft = 1
  ptWidth = 2
  ptHeight = 3
End Enum

Private mregExp As cRegExpression
Private mShape As Shape
Public Position As String

Public Property Let Shape(RHS As Shape)
  Set mShape = RHS
  If Not annotate Then refreshShape RHS
End Property
Private Property Get styleStatement() As String
  styleStatement = mregExp.strParse(mShape.AlternativeText, styleMask)
End Property
Private Property Let styleStatement(RHS As String)
  mShape.AlternativeText = RHS
End Property
Property Get Name() As String
  Name = mShape.OLEFormat.Object.Name
End Property
Property Let Name(RHS As String)
  mShape.OLEFormat.Object.Name = RHS
End Property

Private Property Get positionAttributes() As String()
  positionAttributes = Split(Trim(mregExp.objParse(styleStatement, positionBody)(0).SubMatches(0)), " ")
End Property
Private Function annotate() As Boolean
Dim st As String

  If styleStatement = vbNullString Then
    With mShape.OLEFormat.Object
      styleStatement = getPosition(mShape)
    End With
    annotate = True
  Else
    annotate = False
  End If
  
End Function

Private Function getPosition(s As Shape) As String
Dim retVal As String

  retVal = retVal & startTag & "position: " & s.Top & " " _
                  & s.Left & " " & s.Width & " " & s.Height
  retVal = retVal & endTag
  getPosition = retVal
  
End Function

Private Function refreshShape(s As Shape) As Boolean
Dim sAttributes() As String

  If UBound(positionAttributes) = 0 Then
    refreshShape = False
    Exit Function
  Else
    sAttributes = positionAttributes
    With mShape
      .Top = CDbl(sAttributes(ptTop))
      .Left = CDbl(sAttributes(ptLeft))
      .Width = CDbl(sAttributes(ptWidth))
      .Height = CDbl(sAttributes(ptHeight))
    End With
  End If
  
End Function

Private Sub Class_Initialize()
  Set mregExp = New cRegExpression
End Sub

