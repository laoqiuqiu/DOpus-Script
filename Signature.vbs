option explicit

' Signature
' qiuqiu

' This is a script for Directory Opus.
' See https://www.gpsoft.com.au/DScripts/redirect.asp?page=scripts for development information.

' Called by Directory Opus to initialize the script
Function OnInit(initData)
	With initData
		.name           = "Signature"
		.version        = "1.0"
		.copyright      = "qiuqiu"
		.url            = "http://script.dopus.net/"
		.desc           = "This script returns the first xx bytes in the file."
		.default_enable = true
		.min_version    = "12.0"
		.Config.Lenght  = &H10&
		With .AddColumn
			.name       = "Signature"
			.method     = "OnSignature"
			.label      = "Signature"
			.justify    = "left"
			.autogroup  = true
		End With
	End With
End Function


' Implement the Signature column
Function OnSignature(ColData)
	Dim Signature, HexStr, AscStr, I
	' On Error Resume Next 
	
	If ColData.Item.Is_Dir Or ColData.Item.Size = 0 Then Exit Function

	Set Signature = ColData.Item.Open.Read(Script.Config.Lenght)
	For I = 0 To Signature.Size - 1
		HexStr = HexStr & Right("0" & Hex(Signature(I)), 2) & " "
		AscStr = AscStr & IIf(Signature(I) > 31 And Signature(I) < 127, ChrW(Signature(I)), ".") & " "
	Next
	ColData.Value = "[ " & HexStr & " | " & Left(AscStr, 16) & " ]"
End Function

Function IIf(ByVal Expression, ByVal TruePart, ByVal FalsePart)
    If Expression Then
		If IsObject(TruePart)  Then Set IIf = TruePart  Else IIf = TruePart
	Else
		If IsObject(FalsePart) Then Set IIf = FalsePart Else IIf = FalsePart
	End If
End Function