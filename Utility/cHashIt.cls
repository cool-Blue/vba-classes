VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHashIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function BASE64SHA1(ByVal sTextToHash As String)

    Const cutoff As Integer = 10
    
    Dim asc As Object
    Dim enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Dim bytes() As Byte

    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")

    TextToHash = asc.GetBytes_4(sTextToHash)
    SharedSecretKey = asc.GetBytes_4(sTextToHash)
    enc.key = SharedSecretKey

    bytes = enc.ComputeHash_2((TextToHash))
    BASE64SHA1 = EncodeBase64(bytes)
    BASE64SHA1 = Left(BASE64SHA1, cutoff)

    Set asc = Nothing
    Set enc = Nothing

End Function

Private Function EncodeBase64(ByRef arrData() As Byte) As String

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text

    Set objNode = Nothing
    Set objXML = Nothing

End Function
Private Sub testHash()
  MsgBox BASE64SHA1("avery very very very long name indeedavery very very very long name indeedavery very very very long name indeed")
End Sub


