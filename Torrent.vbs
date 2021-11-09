Class BinaryStream
  Private Stream
  
  Private Sub Class_Initialize   ' Setup Initialize event.
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Type = 1
    Stream.Mode = 3
    Stream.Open
  End Sub
  
  Private Sub Class_Terminate   ' Setup Terminate event.
    Stream.Close
    Set Stream = Nothing
  End Sub
  
  Private Function Array2Bytes(VBS_Array)
    Dim B, MemoryStream
    Set MemoryStream = CreateObject("System.IO.MemoryStream")
    MemoryStream.SetLength(0)
    For Each B In VBS_Array
      MemoryStream.WriteByte CByte(B)
    Next
    Array2Bytes = MemoryStream.ToArray
  End Function
  
  Private Function BytesToArray(Bytes)
    Dim I, L, A()
    If VarType(Bytes) = 8209 Then
      L = UBound(Bytes) - 1
      ReDim A(L)
      If LenB(Bytes) = 0 Then Exit Function
      For I = 0 To L
        A(I) = AscB(MidB(Bytes, I+1, 1))
      Next
      BytesToArray = A
    End If
  End Function
  
  Private Function UTF8Decode(Bytes)
    Dim UTF8 : Set UTF8 = CreateObject("System.Text.UTF8Encoding")
    UTF8Decode = ""
    If VarType(Bytes) = 8209 Then UTF8Decode = UTF8.GetString((Bytes))
    Set UTF8 = Nothing
  End Function
  
  Private Function UTF8Encode(Strings)
    Dim UTF8 : Set UTF8 = CreateObject("System.Text.UTF8Encoding")
    UTF8Encode = Empty
    If VarType(Strings) = 8 Then UTF8Encode = UTF8.GetBytes_4(Strings)
    Set UTF8 = Nothing
  End Function
  
  Private Sub ResetStream
    If Stream.State > 0 Then Stream.Close
    If Stream.State = 0 Then Stream.Open
  End Sub
  
  Public Sub LoadFromFile(FileName)
    ResetStream
    With CreateObject("ADODB.Stream")
      .Type = 1
      .Mode = 3
      .Open
      .LoadFromFile(F1)
      .CopyTo Stream
    End With
    Position = 0
  End Sub
  
  Public Function LoadFromString(Str)
    ResetStream
    Stream.Write UTF8Encode(Str)
    LoadFromString = Size
    Position = 0
  End Function
  
  Public Function ReadByte
    ReadByte = AscB(Stream.Read(1))
  End Function
  
  Public Function ReadBytes(Length)
    ReadBytes = Stream.Read(Length)
  End Function

  Public Function ReadChar
    ReadChar = ChrW(ReadByte)
  End Function
  
  Public Function ReadUTF8String(Length)
    ReadUTF8String = UTF8Decode(Stream.Read(Length))
  End Function
  
  Public Function ReadUntil(Mark)
    Dim C, Result
    C = ReadChar
    Do While C <> Mark
      Result = Result & C
      C = ReadChar
    Loop
    ReadUntil = Result
  End Function
  
  Public Sub Seek(Offset, Mode)
    Select Case Mode
      Case 1 : Stream.Position = Offset
      Case 2 : Stream.Position = Stream.Position + Offset
      Case 3 : Stream.Position = Stream.Size + Offset
    End Select
  End Sub
  
  Public Property Let Position(Offset)
    Stream.Position = Offset
  End Property
  
  Public Property Get Position
    Position = Stream.Position
  End Property
  
  Public Property Get Size
    Size = Stream.Size
  End Property
  
End Class

Sub ArrayAdd(ByRef arr, ByVal Value)
  If IsArray(arr) Then
    On Error Resume Next
    Dim ub :ub = UBound(arr)
    If Err.Number <> 0 Then ub = -1
    ReDim Preserve arr(ub + 1)
    Select Case VarType(Value)
      Case 9, 12, 13
        Set arr(UBound(arr)) = Value
      Case Else
        arr(UBound(arr)) = Value
    End Select
  End If
End Sub

Private Function BytesToHex(Bytes)
  Dim I, L, A()
  If VarType(Bytes) = 8209 Then
    L = UBound(Bytes) - 1
    ReDim A(L)
    If LenB(Bytes) = 0 Then Exit Function
    For I = 0 To L
      A(I) = Right("00" & Hex(AscB(MidB(Bytes, I+1, 1))), 2)
    Next
    BytesToHex = Join(A, "")
  End If
End Function

Function GetArrayDim(ByVal arr)
  Dim i
  If IsArray(arr) Then
    For i = 1 To 60
      On Error Resume Next
      Call UBound(arr, i)
      If Err.Number <> 0 Then
        GetArrayDim = i - 1
        Exit Function
      End If
    Next
    GetArrayDim = i
  Else
    GetArrayDim = Null
  End If
End Function

Function Decode(In_Stream, ByRef Char)
  Select Case Char
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
      Decode = In_Stream.ReadUTF8String(CLng(Char & In_Stream.ReadUntil(":")))
    Case "i"
      Decode = CCur(In_Stream.ReadUntil("e"))
    Case "l"
'      Dim List : Set List = CreateObject("System.Collections.ArrayList")
'      Char = In_Stream.ReadChar
'      Do While Char <> "e"
'        List.Add Decode(Stream, Char)
'        Char = In_stream.ReadChar
'      Loop
'      Set Decode = List
      Dim List()
      Char = In_Stream.ReadChar
      Do While Char <> "e"
        ArrayAdd List, Decode(In_Stream, Char)
        Char = In_stream.ReadChar
      Loop
      Decode = List
    Case "d"
      Dim Key, Dict : Set Dict = CreateObject("scripting.dictionary")
      Char = In_Stream.ReadChar
      Do While Char <> "e"
        Key = Decode(In_Stream, Char)
        Char = In_Stream.ReadChar
        Select Case Key
          Case "ed2k", "md5sum", "filehash", "pieces" ' Non-string value
            Dict.Add Key, In_Stream.ReadBytes(CLng(Char & In_Stream.ReadUntil(":")))
          Case Else
            Dict.Add Key, Decode(In_Stream, Char)
        End Select
        Char = In_Stream.ReadChar
      Loop
      Set Decode = Dict
  End Select
End Function

