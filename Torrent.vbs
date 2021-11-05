Class BinaryStream
  Private Stream
  
  Private Sub Class_Initialize   ' Setup Initialize event.
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Type = 1
    Stream.Mode = 3
    Stream.Open
  End Sub
  
  Private Sub Class_Terminate   ' Setup Terminate event.
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
  
  Public Sub LoadFromFile(FileName)
    Stream.LoadFromFile(FileName)
  End Sub
  
  Public Function LoadFromString(Str)
    Stream.Write UTF8Encode(Str)
    LoadFromString = Size
    Position = 0
  End Function
  
  Public Function ReadByte
    ReadByte = AscB(Stream.Read(1))
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

Function Decode(In_Stream, ByRef Char)
  Select Case Char
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
      Decode = In_Stream.ReadUTF8String(CLng(Char & In_Stream.ReadUntil(":")))
    Case "i"
      Decode = CCur(In_Stream.ReadUntil("e"))
    Case "l" ' list(index)(0) = list_item
      Dim List : Set List = CreateObject("System.Collections.ArrayList")
      Char = In_Stream.ReadChar
      Do While Char <> "e"
        List.Add Decode(Stream, Char)
        Char = In_stream.ReadChar
      Loop
      Set Decode = List
    Case "d"
      Dim Dict : Set Dict = CreateObject("scripting.dictionary")
      Dim Key
      Char = In_Stream.ReadChar
      Do While Char <> "e"
        Key = Decode(Stream, Char)
        Char = In_Stream.ReadChar
        Dict.Add Key, Decode(Stream, Char)
        Char = In_Stream.ReadChar
      Loop
      Set Decode = Dict
  End Select
End Function

