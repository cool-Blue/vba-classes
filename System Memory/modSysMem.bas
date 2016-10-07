Attribute VB_Name = "modSysMem"
Option Explicit
Type pointerChainType
    objName As String
    lptr As LongPtr
    plptr As LongPtr
    pplptr As LongPtr
End Type
Const ErrNotDivisible As Long = vbObjectError + 1
Const Err1Message As String = "Mem_ReadHex_Words: Length is not divisible by wordLength"
Const ErrZeroPointer As Long = vbObjectError + 2
Const Err2Message As String = "Mem_ReadHex_Words: wordLength Exceeds LongPtr size."
Const Err3Number As Long = vbObjectError + 3
Const Err3Message As String = "Mem_ReadHex_Words: Illegal Pointer Value: zero pointer not allowed."

'
' http://social.msdn.microsoft.com/Forums/eu/exceldev/thread/e3aefd82-ec6a-49c7-9fbf-5d57d8ef65ca
'

Public Enum enInit
    memGetDelta = False
    memGetReference = True
End Enum
Type PROCESS_MEMORY_COUNTERS
   cb                         As Long
   PageFaultCount             As Long
   PeakWorkingSetSize         As Long
   WorkingSetSize             As Long
   QuotaPeakPagedPoolUsage    As Long
   QuotaPagedPoolUsage        As Long
   QuotaPeakNonPagedPoolUsage As Long
   QuotaNonPagedPoolUsage     As Long
   PagefileUsage              As Long
   PeakPagefileUsage          As Long
End Type

#If Win64 Then
    Public Const PTR_LENGTH As Long = 8
#Else
    Public Const PTR_LENGTH As Long = 4
#End If
 
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long

Public Declare PtrSafe Sub Mem_Copy Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)
Sub crashMem_Copy()
Dim lptr As LongPtr
'    Mem_Copy lptr, ByVal 0, 4      'crashes excel
'    Mem_Copy lptr, ByVal lptr, 4   'crashes excel
    Mem_Copy lptr, lptr, 4            'is fine
End Sub

  
Public Function GetCurrentProcessMemory(Optional bInit As enInit = memGetDelta, Optional ByRef savedRef As Long = -1) As String

  Dim lngCBSize2           As Long
  Dim lngModules(1 To 200) As Long
  Dim lngReturn            As Long
  Dim lngHwndProcess       As Long
  Dim pmc                  As PROCESS_MEMORY_COUNTERS
  Dim lRet                 As Long
  Dim MemDelta             As Long
  Dim sRet As String
  Static MemUsed           As Long
  

  If bInit Then MemUsed = 0

  'Get a handle to the Process and Open
  lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, GetCurrentProcessId)

  If lngHwndProcess <> 0 Then

      'Get an array of the module handles for the specified process
      lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCBSize2)

      'If the Module Array is retrieved, Get the ModuleFileName
      If lngReturn <> 0 Then
         
         'Get the Site of the Memory Structure
          pmc.cb = LenB(pmc)

          lRet = GetProcessMemoryInfo(lngHwndProcess, pmc, pmc.cb)

          If bInit Then
              sRet = "WorkingSetSize  " & Format(Fix(pmc.WorkingSetSize), "###,###,###") & "  ( Reference )"
              MemUsed = pmc.WorkingSetSize
              If savedRef < 0 Then savedRef = MemUsed
          Else
              MemDelta = pmc.WorkingSetSize - max(MemUsed, savedRef)
              If MemDelta > 0 Then
                 sRet = "WorkingSetSize  " & Format(Fix(pmc.WorkingSetSize), "###,###,###") & _
                                         "  ( " & Format(Fix(MemDelta), "###,###,###") & " )"
              Else
                 sRet = "WorkingSetSize  " & Format(Fix(pmc.WorkingSetSize), "###,###,###") & "  ( All Cleared )"
              End If
          End If
          
      End If

  End If

  GetCurrentProcessMemory = sRet
  'Close the handle to this process
  lngReturn = CloseHandle(lngHwndProcess)

End Function
Public Sub testMem()
Attribute testMem.VB_ProcData.VB_Invoke_Func = "M\n14"
    MsgBox GetCurrentProcessMemory
End Sub
' Platform-independent method to return the full zero-padded
' hexadecimal representation of a pointer value
Function HexPtr(ByVal Ptr As LongPtr) As String
'Debug.Print Ptr
    HexPtr = Hex$(Ptr)
    HexPtr = String$((PTR_LENGTH * 2) - Len(HexPtr), "0") & HexPtr
End Function
Public Function Mem_Read_WordX(ByVal Ptr As LongPtr, Optional prefix As String = "0x") As String
    Mem_Read_WordX = prefix & HexPtr(Mem_Read_Word(Ptr))
End Function

Public Function Mem_Read_Word(ByVal Ptr As LongPtr) As LongPtr
Const ErrSource As String = "Mem_Read_Word"
Dim lptr As LongPtr
    If Ptr = 0 Then
        Err.Raise Number:=Err3Number, Source:=ErrSource, Description:=Err3Message
        Exit Function
    End If
    Mem_Copy ByVal VarPtr(lptr), ByVal Ptr, PTR_LENGTH
    Mem_Read_Word = lptr
End Function
Public Function Mem_ReadHex_Words(ByVal Ptr As LongPtr, ByVal Length As Long, _
                                    Optional wordLength As Long = PTR_LENGTH) As String
Const ErrSource As String = "Mem_ReadHex_Words"
Dim wordBuffer As LongPtr, strBytes As String
Dim WordCount As Long, w As Long
    
    If Ptr = 0 Then
        Err.Raise Number:=Err3Number, Source:=ErrSource, Description:=Err3Message
        Exit Function
    End If
    If wordLength > PTR_LENGTH Then
        Err.Raise Number:=ErrZeroPointer, Source:=ErrSource, Description:=Err2Message
        Exit Function
    End If
    WordCount = Length \ wordLength
    If WordCount <> Length / wordLength Then
        Err.Raise Number:=ErrNotDivisible, Source:=ErrSource, Description:=Err1Message
        Exit Function
    End If
    
    For w = 0 To WordCount - 1
        Mem_Copy wordBuffer, ByVal Ptr + w * wordLength, wordLength
        strBytes = strBytes & Right(HexPtr(CLng(wordBuffer)), wordLength * 2) & " "
    Next w
    Mem_ReadHex_Words = Left(strBytes, Len(strBytes) - 1) 'Discard trailing space
    
End Function

Sub testMemReadHexWords()
Const byteCount As Long = 16
Const wordLen As Long = PTR_LENGTH
Dim lptr As LongPtr, i As Long, WordCount As Long
Dim words() As String, str As String
    lptr = ObjPtr(Worksheets(1))
    Debug.Print Mem_ReadHex_Words(lptr, byteCount, wordLen)
    str = Mem_ReadHex(lptr, byteCount)
    WordCount = byteCount / wordLen
    ReDim words(WordCount - 1)
    For i = 0 To WordCount - 1
        words(i) = Mid(str, i * wordLen * 2 + 1, wordLen * 2)
    Next i
    Debug.Print Join(words, " ")
End Sub

Public Function Mem_ReadHex(ByVal Ptr As LongPtr, ByVal Length As Long, Optional Reverse As Boolean = False) As String
Dim bBuffer() As Byte, strBytes() As String, i As Long, j As Long, UB As Long, b As Byte
Dim forStart As Long, forEnd As Long, forStep As Long
    
    UB = Length - 1
    ReDim bBuffer(UB)
    ReDim strBytes(UB)
    
    If Reverse Then
        forStart = UB: forEnd = 0: forStep = -1
    Else
        forStart = 0: forEnd = UB: forStep = 1
    End If
    Mem_Copy bBuffer(0), ByVal Ptr, Length
    j = 0
    For i = forStart To forEnd Step forStep
        b = bBuffer(i)
        strBytes(j) = IIf(b < 16, "0", "") & Hex$(b)
        j = j + 1
    Next
    Mem_ReadHex = Join(strBytes, "")
End Function

Public Function ptrValue(lptr As LongPtr) As LongPtr
'Returns the address of the referenced Object
Const ErrSource As String = "ptrValue"
Dim dest As LongPtr
    If lptr = 0 Then
        Err.Raise Number:=Err3Number, Source:=ErrSource, Description:=Err3Message
        Exit Function
    End If
    Mem_Copy dest, ByVal lptr, PTR_LENGTH
    ptrValue = dest
End Function
Public Function objAddress(obj As Object) As LongPtr
'Returns the address of the referenced Object
Dim dest As LongPtr
    Mem_Copy dest, obj, PTR_LENGTH
    objAddress = dest
End Function

'Functions:
'==========
'   memBlock
'   FUNC:   Returns a string, containing a printout of a memory
'           block specified
'   INPUT:  StrtMem As Long  The memory location from where the
'                            printout begins.
'   Length As Long           The length of the memory to be printed
'                            out, in bytes.
'   [Title] As String        Optional. A heading to be printed of
'                            the printout. If blank ("", default),
'                            then no heading is printed.
'   [Reverse] As Boolean     Optional. If TRUE the memory printout
'                            is from the highest byte to the lowest.
'                            Default is FALSE.
'   RETURN: As String        The format of the printout is as follows:
'                            [Heading]
'                            =========
'                            0  ([StrtMem+0]) >> Value0_Char
'                            Value0_Hex (Value0_Dec) [Value0_Bin]
'                            1  ([StrtMem+1]) >> Value1_Char
'                            Value1_Hex(Value1_Dec) [Value1_Bin]
'                            2  ([StrtMem+2]) >> Value2_Char
'                            Value2_Hex (Value2_Dec) [Value2_Bin]
'                            " >> " " " "
'                            " >> " " " "
'                            L  ([StrtMem+L]) >> ValueL_Char
'                            ValueL_Hex (ValueL_Dec) [ValueL_Bin]
'


'FUNC:  Returns a string, containing a printout of a memory block specified
Public Function memBlock(ByVal StrtMem As Long, _
                         ByVal Length As Long, _
                         Optional Title As String = "", _
                         Optional Reverse As Boolean = False) As String

   Dim i                 As Long
   Dim strUnderline      As String
   Dim lng1              As Byte
   Dim ptrToLong         As Long

   Dim lngForStart       As Long
   Dim lngForFinish      As Long
   Dim lngForStep        As Long

   Dim strTempReturn     As String

   strTempReturn = ""

   'Setup the for loop variables
   If Not (Reverse) Then
      lngForStart = 0
      lngForFinish = Length - 1
      lngForStep = 1
   Else
      lngForStart = Length - 1
      lngForFinish = 0
      lngForStep = -1
   End If

   lng1 = 0
   ptrToLong = VarPtr(lng1)

   'Print the heading
   If Title <> "" Then
      strTempReturn = strTempReturn & Title & ":" & vbNewLine
      strUnderline = String(Len(Title) + 1, "-")
      strTempReturn = strTempReturn & strUnderline & vbNewLine
   End If

'   For i = Length - 1 To 0 Step -1
   For i = lngForStart To lngForFinish Step lngForStep
      Mem_Copy ByVal ptrToLong, ByVal (StrtMem + i), 1
      'Write the relative, and absolute memory addresses
      If (i < 10) Then
         strTempReturn = strTempReturn & i & " (" & Hex(StrtMem + i) & ") >> "
      Else
         strTempReturn = strTempReturn & i & " (" & Hex(StrtMem + i) & ") >> "
      End If

      'Write the content as a character
      If (lng1 >= 32) And (lng1 <= 255) Then
         strTempReturn = strTempReturn & Chr(lng1)
      Else
         strTempReturn = strTempReturn & "."
      End If

      'Write the content as a hexidecimal number
      If Len(Hex(lng1)) = 1 Then
         strTempReturn = strTempReturn & " " & Hex(lng1) & "h "
      Else
         strTempReturn = strTempReturn & " " & Hex(lng1) & "h "
      End If

      'Write the content as a decimal number
      If (lng1 < 10) Then
         strTempReturn = strTempReturn & " (" & lng1 & ") "
      ElseIf (lng1 < 100) Then
         strTempReturn = strTempReturn & " (" & lng1 & ") "
      Else
         strTempReturn = strTempReturn & "(" & lng1 & ") "
      End If

      'Write the content as a binary number
      strTempReturn = strTempReturn & "[" & Dec2Bin(lng1) & "]"

      'Write a new line character
     strTempReturn = strTempReturn & vbNewLine
   Next i
   memBlock = strTempReturn & vbNewLine
End Function


Private Function Dec2Bin(ByVal intDec As Byte) As String
   Dim strHex          As String
   Dim strTempReturn   As String
   Dim strTempNibble   As String * 4
   Dim i               As Integer

   'Initialise the return variable
   strTempReturn = String(8, "*")

   'Get the hexidecimal value of the binary number
   strHex = Hex(intDec)

   'Test if strTempHex is 1 character long
   If (Len(strHex) = 1) Then
      strHex = "0" & strHex
   End If

   'Convert the hexidecimal number to binary

   For i = 1 To 2
      Select Case (Mid$(strHex, i, 1))
      Case "0"
         strTempNibble = "0000"
      Case "1"
         strTempNibble = "0001"
      Case "2"
         strTempNibble = "0010"
      Case "3"
         strTempNibble = "0011"
      Case "4"
         strTempNibble = "0100"
      Case "5"
         strTempNibble = "0101"
      Case "6"
         strTempNibble = "0110"
      Case "7"
         strTempNibble = "0111"
      Case "8"
         strTempNibble = "1000"
      Case "9"
         strTempNibble = "1001"
      Case "A"
         strTempNibble = "1010"
      Case "B"
         strTempNibble = "1011"
      Case "C"
         strTempNibble = "1100"
      Case "D"
         strTempNibble = "1101"
      Case "E"
         strTempNibble = "1110"
      Case "F"
         strTempNibble = "1111"
      End Select

      If i = 1 Then
         Mid(strTempReturn, 1, 4) = strTempNibble
      Else
         Mid(strTempReturn, 5, 4) = strTempNibble
      End If
   Next i

   'Return the value in binary
   Dec2Bin = strTempReturn
End Function
Sub formatCols()
Const colwidth As Long = 10
Dim message As clStingsToColumns

    Set message = New clStingsToColumns
    With message
        .tabs = Array(colwidth, colwidth, colwidth, colwidth)
        .indents = 1
        Debug.Print .Format("col1", "col2", "col3", "col4")
    End With
    
End Sub
Sub formatCols1()
Const colwidth As Long = 10
Dim message As clStingsToColumns

    Set message = New clStingsToColumns
    With message
        .setCols colwidth, colwidth, colwidth, colwidth
        .setIndents 1, 2, 1
        Debug.Print .Format("col1", "col2", "col3", "col4")
    End With
    
End Sub


