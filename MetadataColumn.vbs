option explicit

' MetadataColumn
' qiuqiu

' Called by Directory Opus to initialize the script
Function OnInit(initData)
	initData.name = "MetadataColumn"
	initData.version = "1.0"
	initData.copyright = "qiuqiu"
	initData.url = "http://script.dopus.net/"
	initData.desc = "Parse data from Metadata"
	initData.default_enable = true
	initData.min_version = "12.0"

	Dim col

	Set col = initData.AddColumn
	col.name = "Platform"
	col.method = "OnParse"
	col.label = "Platform"
	col.justify = "left"
	col.autogroup = true
	col.multicol = true
	col.justify = "center"
	
	Set col = initData.AddColumn
	col.name = "signature"
	col.method = "OnParse"
	col.label = "signature"
	col.justify = "left"
	col.autogroup = true
	col.multicol = true
	col.justify = "center"

	Set col = initData.AddColumn
	col.name = "IsDotNET"
	col.method = "OnParse"
	col.label = "Is .NET"
	col.justify = "left"
	col.autogroup = true
	col.multicol = true
	col.justify = "center"
End Function


Function IIf(Exp, RTrue, RFalse)
 If Exp Then
  IIf = RTrue
 Else
  IIf = RFalse
 End If
End Function

' Implement the Platform column
Function OnParse(ColData)
	Dim meatdata, mdarr
	
	if ColData.item.is_dir then exit function
	On Error Resume Next
	meatdata = Replace(LCase(ColData.item.metadata.other.autodesc), " ", "")
	'dopus.output ColData.item.name & " - " & meatdata

	if len(meatdata) = 0 then exit function

	mdarr = split(meatdata,",")
	select case mdarr(0)
		case "i386", "amd64"
			ColData.Columns("Platform").Value  = IIf(mdarr(0) = "i386", "Win32", "Win64")

			If (mdarr(1) = ".net") then 
				ColData.Columns("IsDotNET").Value = "Yes"
				ColData.Columns("signature").Value = mdarr(2)
			else
				ColData.Columns("signature").Value = mdarr(1)
			End If
	End Select
End Function
