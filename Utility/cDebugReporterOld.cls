VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDebugReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/ Maintains a global call stack
'/ Requires
'/ Public const gcDebugMode as boolean = ???
'/ Public Const gcContextCol As Long = -1   "-1" to dissable
'/ Public gCallDepth as Long
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private mLevel As Long
Private mFirstReport As Boolean
Private mCaller As String
Private mContext As String
Private mContextCol As Long
Private mMessage As String
Private mMessageCol As Long

Private Sub Class_Initialize()
    gCallDepth = gCallDepth + 1
    mFirstReport = True
    mContextCol = gcContextCol
    mMessageCol = -1
End Sub
Public Sub Report(Optional d As Double = 0, Optional caller As String, Optional Context As String, _
                            Optional contextCol As Long = -1, Optional Message As String, _
                            Optional messageCol As Long = -1)

    If gcDebugMode Then
    
    Dim newLine As eNewLine
    Dim myContext As String
    
        If Len(caller) <> 0 Then mCaller = caller
        If contextCol > 0 Then mContextCol = contextCol
        If Len(Context) = 0 Then
            If mFirstReport Then
                myContext = "START"
            Else
                myContext = mContext
            End If
        Else
            mContext = Context
            myContext = mContext
        End If
        newLine = IIf(mFirstReport, Before, No)
        Debug.Print timeStamp(d:=d, newLine:=newLine, _
                                Indent:=gCallDepth, _
                                caller:=IIf(mFirstReport, mCaller, String(Len(mCaller), " ")), _
                                Context:=myContext, _
                                contextCol:=mContextCol, _
                                Message:=Message, _
                                messageCol:=max(messageCol, mMessageCol))
    
        If mFirstReport Then mCaller = caller
        mContextCol = contextCol
        mFirstReport = False
    End If
    
End Sub
Property Let Context(RHS As String)
    mContext = RHS
End Property
Property Let Message(RHS As String)
    mMessage = RHS
End Property
Private Sub Class_Terminate()
    If gcDebugMode Then
        Debug.Print timeStamp(newLine:=After, _
                                Indent:=gCallDepth, _
                                caller:=mCaller, _
                                Context:="End", _
                                contextCol:=mContextCol)
    End If
    gCallDepth = gCallDepth - 1
End Sub

