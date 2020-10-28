Option Explicit

' CopyContent
' (c) 2020 qiuqiu

' Arguments
' APPEND    (Switch)   Optional, Append to clipboard.
' FILE      (multiple) Optional, The full path to one or more files, Multiple files are separated by spaces.
' FILEINFO  (Switch)   Optional, include file information.
' PATTERN   (keyword)  Optional, Specify the wildcard pattern which files must match.
' REGEXP    (Switch)   Optional, Enables regular expression mode.
' TRIM      (keyword)   Optional, Without leading spaces (left), trailing spaces (right), or both leading and trailing spaces (all).

' Examples:
' Send the contents of the specified file to the clipboard.
' CopyContent c:\test.txt
'
' Send the contents of the file with the extension txt in the specified folder to the clipboard.
' CopyContent c:\temp PATTERN *.(txt|vbs) FILEINFO TRIM
'
' Send file content to clipboard for all selected TXT files only.
' CopyContent PATTERN *.TXT 

' History:
' 2020-10-28 First version


' Called by Directory Opus to initialize the script
Function OnInit(initData)
    With InitData
        .Name           = "Copy File Content"
        .Version        = "1.0"
        .Copyright      = "(c) 2020 qiuqiu"
        .Url            = "http://script.dopus.net/"
        .Desc           = Dopus.Strings.Get("desc")
        .Default_Enable = True
        .Min_Version    = "12.0" ' Used a feature included in 12.20.7, I did not test it in DOpus earlier than this version.
        .group          = Dopus.Strings.Get("group")
        
        With .AddCommand
            .Name     = "CopyContent"
            .Method   = "OnCopyContent"
            .Desc     = Dopus.Strings.Get("desc")
            .Label    = "CopyContent"
            .Template = "APPEND/S/O,FILEINFO/S/O,FILE/M,PATTERN/K/O,REGEXP/S/O,TRIM/K/O[all,left,right]"
            .Hide     = False
            .Icon     = "copy"
        End With
    End With
End Function

'Implement the CopyContent command
Function OnCopyContent(CmdData)
    Dim CilpText, Result, Text, FileInfo, Files, Flags, i, kTrim
    Set Files = DOpus.Create.Vector
    
    If CmdData.Func.Args.got_arg.trim Then
        If len(CmdData.Func.Args.trim) Then
            kTrim = Split(LCase(CmdData.Func.Args.trim), ",")(0)	
        Else
            kTrim = "all"
        End If
    End If
    
    If CmdData.Func.Args.got_arg.file Then
        Files.Assign CmdData.Func.Args.file
    ElseIf CmdData.Func.Command.Filecount > 0 Then
        Files.Assign CmdData.Func.Command.Files
    End If
    
    If Files.Empty Then Exit Function
    CmdData.Func.Command.ClearFiles
    For Each i In Files
        CmdData.Func.Command.AddFiles GetFiles(i, False)
    Next ' i
    
    Files.Assign CmdData.Func.Command.Files
    If CmdData.Func.Args.got_arg.Pattern Then
        CmdData.Func.Command.ClearFiles
        CmdData.Func.Command.AddFiles FilterFiles(Files, CmdData.Func.Args.Pattern, CmdData.Func.Args.got_arg.RegExp)
    End If
    
    For Each i In CmdData.Func.Command.Files
        Select Case kTrim
            Case "all"   : Text = Trim(ReadText(i))
            Case "left"  : Text = TrimL(ReadText(i))
            Case "right" : Text = TrimR(ReadText(i))
            Case Else    : Text = ReadText(i)
        End Select
        
        If CmdData.Func.args.got_arg.fileinfo Then FileInfo = "[" & i & "]" & vbNewLine
        Result = Result & vbNewLine & FileInfo & Text
    Next 'i
    
    Result = Trim(Result)
    
    If Len(Result) = 0 Then Exit Function
    If CmdData.Func.args.got_arg.append Then
        If DOpus.GetClipFormat = "text" Then CilpText = DOpus.GetClip
        DOpus.SetClip CilpText & vbNewLine & Result
    Else
        DOpus.SetClip Result
    End If
    
End Function

'Returns a copy of a string without leading spaces.
Function TrimL(ByVal str)
    Do While True
        Select Case left(str, 1)
            Case vbCR, vbLF, vbTab, vbVerticalTab, " "
            str = Right(str, Len(str) - 1)
            Case Else
            Exit Do
        End Select
    Loop
    TrimL = str
End Function

'Returns a copy of a string without trailing spaces (TrimR).
Function TrimR(ByVal str)
    Do While True
        Select Case Right(str, 1)
            Case vbCR, vbLF, vbTab, vbVerticalTab, " "
            str = Left(str, Len(str) - 1)
            Case Else
            Exit Do
        End Select
    Loop
    TrimR = str
End Function

'Returns a copy of a string without leading and trailing spaces.
Function [Trim](ByVal str)
    Trim = TrimL(TrimR(str))
End Function

''' <summary>Returns a Vector object that lets you enumerate the contents of the specified folder.</summary>
''' <param name="strPath" type="string">Path string</param>
''' <param name="blnRecurse" type="Boolean">Recursively enumerate the folder.</param>
Function GetFiles(ByVal strPath, ByVal blnRecurse)
    Dim Flags, f
    Set GetFiles = DOpus.Create.Vector
    Select Case DOpus.FSUtil.GetType(strPath) 
        Case "dir"
        If blnRecurse Then Flags = "r" Else Flags = Empty
        For Each f In DOpus.FSUtil.ReadDir(strPath, Flags).Next(-1)
            If Not f.is_dir Then GetFiles.push_back f
        Next ' f 
        Case "file"
        GetFiles.push_back DOpus.FSUtil.GetItem(strPath)
    End Select
End Function

''' <summary>Returns a Vector object of the wildcard pattern which files must match.</summary>
''' <param name="Files" type="Vector">collection:Item</param>
''' <param name="strPattern" type="string">Specify the wildcard pattern which files must match.</param>
''' <param name="blnREGEXP" type="Boolean">Enables regular expression mode.</param>
Function FilterFiles(ByVal Files, ByVal strPattern, ByVal blnREGEXP)
    Dim Wild, Flags, f
    
    If blnREGEXP Then Flags = "fr" Else Flags = "f"
    
    Set Wild = DOpus.FSUtil.NewWild(strPattern, Flags)
    Set FilterFiles = DOpus.Create.Vector
    
    For Each f In Files
        If Wild.Match(f.name)Then
            FilterFiles.push_back f
        End If
    Next
End Function

Sub AppendVector(ByRef Vector1, ByVal Vector2)
    Dim i
    For Each i In Vector2
        Vector1.push_back i
    Next
End Sub

''' <summary>check byte array is utf-8 string</summary>
Function CheckUTF8(ByRef Byte_Array)
    ' UTF8 Valid sequences
    ' 0xxxxxxx  ASCII
    ' 110xxxxx 10xxxxxx  2-byte
    ' 1110xxxx 10xxxxxx 10xxxxxx  3-byte
    ' 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx  4-byte
    ' Width in UTF8
    ' Decimal		Width
    ' 0-127		    1 byte
    ' 194-223		2 bytes
    ' 224-239		3 bytes
    ' 240-244		4 bytes
    '
    ' Subsequent chars are in the range 128-191
    Dim pos, length, ch, more_chars, only_saw_ascii_range 
    only_saw_ascii_range = True
    pos = 0
    length = UBound(Byte_Array)
    Do While pos < length 
        
        ch = Byte_Array(pos)
        pos = pos + 1
        
        If ch = 0 Then
            CheckUTF8 = "None"            
            Exit Function
        ElseIf ch <= 127 Then		
            more_chars = 0		             ' 1 byte
        ElseIf ch >= 194 And ch <= 223 Then		
            more_chars = 1		             ' 2 Byte
        ElseIf ch >= 224 And ch <= 239 Then
            more_chars = 2	                 ' 3 Byte
        ElseIf ch >= 240 And ch <= 244 Then		
            more_chars = 3		             ' 4 Byte
        Else		
            CheckUTF8 = "None"	             ' Not utf8
            Exit Function
        End If
        
        ' Check secondary chars are in range if we are expecting any
        Do While more_chars And pos < length		
            only_saw_ascii_range = False	' Seen non-ascii chars now
            
            ch = Byte_Array(pos)
            pos = pos + 1
            If ch < 128 Or ch > 191 Then
                CheckUTF8 = "None"			' Not utf8                
                Exit Function
            End If
            
            more_chars = more_chars - 1
        Loop
        
    Loop
    
    ' If we get to here then only valid UTF-8 sequences have been processed
    ' If we only saw chars in the range 0-127 then we can't assume UTF8 (the caller will need to decide)
    If only_saw_ascii_range Then
        CheckUTF8 = "ASCII"
    Else        
        CheckUTF8 = "UTF-8"
    End If
End Function

''' <summary>Use ADODB.Stream to read text files, able to recognize most document encodings.</summary>
Function ADOReadText(ByVal FileName)
    With CreateObject("ADODB.Stream")
        ' adTypeText  = 2, adTypeBinary = 1
        .Type = 2 
        .Open
        .Charset = "_autodetect_all"
        .LoadFromFile FileName
        ADOReadText = .ReadText
        .Close
    End With
    'Remove ZERO WIDTH NO-BREAK SPACE
    If ((AscW(ADOReadText) And &HFFFF&) = &HFEFF&) Then ADOReadText = Mid(ADOReadText, 2)
End Function

Function ReadText(ByVal File)
    Dim Blob, StringTools, Text, Encoding
    
    Set StringTools = DOpus.Create.StringTools
    With DOpus.FSUtil.OpenFile(File)
        If .Error = 0 Then Set Blob = .Read
        .Close
    End With
    If Blob.Size Then
        Encoding = CheckUTF8(Blob.ToVBArray)
        If (Encoding = "UTF-8") Or (Encoding = "ASCII") Then
            Text = StringTools.Decode(Blob, "UTF-8")
        Else
            On Error Resume Next
            Text = StringTools.Decode(Blob, "auto") ' 12.20.7, The scripting StringTools object's Encode and Decode methods can now convert to and from raw UTF-16 data, including support for both Big Endian and Little Endian, and optional Byte Order Marks.
            If Err.Number = 5 Then Text = ADOReadText(File)
        End If
        
        Set StringTools = Nothing : Set Blob = Nothing
        ReadText = Text
    End If
End Function

==SCRIPT RESOURCES
<resources>
    <resource type = "Strings">
        <Strings lang = "english">
            <string id = "desc"     text = "Send text file content to clipboard." />
			<string id = "group"    text = "File Command" />
		</Strings>
		<Strings lang = "chs">
            <string id = "desc"     text = "将文本文件内容发送到剪贴板。" />
			<string id = "group"    text = "文件命令" />
        </Strings>
    </resource>
</resources>
