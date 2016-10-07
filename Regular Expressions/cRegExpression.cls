VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRegExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'/////////////////////////////////////////////////////////////////////////////////////////////////////
'/Reference: Microsoft VBScript Regular Expressions
'/////////////////////////////////////////////////////////////////////////////////////////////////////
Private regExpression As RegExp

Public isGlobal As Boolean
Public isCaseSensitive As Boolean
Public isMultiLine As Boolean

Private matchSet As MatchCollection

Private Sub Class_Initialize()
  Set regExpression = New RegExp
  
' Set environmental context
  isGlobal = True
  isCaseSensitive = True
  isMultiLine = False
  
End Sub
Function strParseRegExp(ByVal whichString As String, _
                        ByVal matchPattern As String, _
                        Optional ByVal occurence As Long = 1, _
                        Optional ByVal replaceWith As String = vbNullString) As String

  With regExpression
  
'   Non-breaking spaces cause problems so get rid of them
    .Pattern = Chr(160): whichString = .Replace(whichString, " ")

'   Set the search pattern
    .Pattern = matchPattern
    
'   Execute
    If .test(whichString) Then
      If replaceWith <> vbNullString Then
        strParseRegExp = .Replace(whichString, replaceWith)
      Else
        strParseRegExp = .Execute(whichString)(occurence - 1).SubMatches(0)
      End If
    End If
    Debug.Print .Pattern
  End With
End Function

Function objParseRegExp(ByVal whichString As String, _
                        ByVal matchPattern As String) As MatchCollection

  With regExpression
  
'   Non-breaking spaces cause problems so get rid of them
    .Pattern = Chr(160): whichString = .Replace(whichString, " ")

'   Set environmental context
    .IgnoreCase = Not isCaseSensitive
    .Global = isGlobal
    .MultiLine = isMultiLine

'   Set the search pattern
    .Pattern = matchPattern
    
'   Execute
    If .test(whichString) Then
        Set objParseRegExp = .Execute(whichString)
    End If
  End With

End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////////
'/Example usage
'/////////////////////////////////////////////////////////////////////////////////////////////////////
Sub ParseCityStateZip(Rng As Range)

  Dim c As Range
  Dim regExpression As cRegExpression
  Dim patterns As Variant, i As Long
  
'   Match:  One or more (+) letters (\w) or spaces (\s) preceeding a comma followed by
'           zero or more (*) non-linebreak characters (.).
'   Return: City
    Const City As String = "([\w\s]+),.*"
    
'   Match:  Zero or more non-linebreak characters (.*) preceeding a comma (\,) followed by
'           one or more spaces (\s+), followed by a valid state abreviation, or S followed by a *,
'           followed by zero or more non-linebreak characters.
'   Return: State Code
    Const State As String = ".*\,\s+(A[AELKPSZ]|C[AOT]|D[EC]|FL|G[AU]|HI|I[ADLN]|" _
                           & "K[SY]|LA|M[AEDINOST]|N[CDEHJMY]|O[HKR]|P[AR]|RI|S[CD]|" _
                           & "T[NX]|V[AT]|W[AIVY]).*"
              
'   Match:  Zero or more non-linebreak characters followed by five digits, followed by
'           zero or one expression consisting of a dash followed by four digits at the end of the expression.
'   Return: ZIP/Post Code
    Const ZipCode As String = ".*(\d{5}(\-\d{4})?)$"
    
    patterns = Array(City, State, ZipCode)
    Set regExpression = New cRegExpression
        
      For Each c In Rng
      
        For i = LBound(patterns) To UBound(patterns)
          With regExpression
            c.Offset(0, i + 1).Value = .strParseRegExp(c.Text, patterns(i))
          End With
        Next i
        
      Next c
     
Exit Sub
    
End Sub

