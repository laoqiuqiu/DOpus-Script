option explicit

' Print
' (c) 2019 qiuqiu

' This is a script for Directory Opus.
' See https://www.gpsoft.com.au/DScripts/redirect.asp?page=scripts for development information.
' Called by Directory Opus to initialize the script
Const clRed    = "#FF0000"
Const clBule   = "#0000FF"
Const clGreen  = "#00B200"
Const clBlack  = "#000000"
Const clWrite  = "#FFFFFF"
Const clPurple = "#8000FF"
Const O_text   = "<font color={0}>{1}: {2}</font>"

Function OnInit(initData)
	With initData
		.name           = "Output debug messages."
		.version        = "1.0"
		.copyright      = "(c) 2019 qiuqiu"
		.desc           = "Output debug messages in command line."
		.default_enable = true
		.min_version    = "11.0"

		Dim cmd
		Set cmd = .AddCommand
		With cmd
			.name     = "Debug"
			.method   = "OnPrint"
			.desc     = "Output debug messages in command line."
			.label    = "Debug"
			.template = "TYPE/K,TEXT/K,COLOR/O"
			.hide     = false
			.icon     = "Logs"
		End With
	End With
End Function


' Implement the Print command
Function OnPrint(scriptCmdData)
	Dim Cmd_Args, Arg_Type, Out_Type, Out_Text, Out_Color
	Set Cmd_Args = scriptCmdData.Func.Args.got_arg
	If Cmd_Args.Type And Cmd_Args.Text Then
		Arg_Type = scriptCmdData.Func.Args.Type
		Out_Text = scriptCmdData.Func.Args.Text
		Select Case UCase(Arg_Type)
			Case "DEBUG", "D"
				Out_Color = clBule
				Out_Type = "Debug"
			Case "ERROR", "E"
				Out_Color = clRed
				Out_Type = "Error"
			Case "INFO", "I"
				Out_Color = clGreen
				Out_Type = "Info"
			Case Else
				If Cmd_Args.Color Then
					Out_Color = scriptCmdData.Func.Args.Color
				Else
					Out_Color = clBlack
				End If
				Out_Type = Arg_Type
		End Select
		DOpus.Output StringFormat(O_text, array(Out_Color, Out_Type, Out_Text))
	End If
End Function

Function StringFormat(ByVal SourceString, Arguments)
   Dim objRegEx  ' regular expression object
   Dim objMatch  ' regular expression match object
   Dim strReturn ' the string that will be returned

   Set objRegEx = New RegExp
   objRegEx.Global = True
   objRegEx.Pattern = "(\{)(\d)(\})"

   strReturn = SourceString
   For Each objMatch In objRegEx.Execute(SourceString)
      strReturn = Replace(strReturn, objMatch.Value, Arguments(CInt(objMatch.SubMatches(1))))
   Next

   StringFormat = strReturn
End Function