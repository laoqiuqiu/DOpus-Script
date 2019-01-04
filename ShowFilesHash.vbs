option explicit

' ShowFilesHash
' qiuqiu

Const SCRIPT_DEBUG = False

Sub Output(ByVal Msg)
	If IsDebug Then	DOpus.Output Msg
End Sub

Function IsDebug()
	On Error Resume Next
	IsDebug = Script.Config.Debug Or SCRIPT_DEBUG
	If Err.Number <> 0 Then
		Err.Clear
		IsDebug = SCRIPT_DEBUG
	End If
	If Err.Number <> 0 Then	IsDebug = False : Err.Clear
End Function
		
Function OnInit(initData)
	With initData
	.name           = "ShowFilesHash"
	.version        = "0.1"
	.copyright      = "qiuqiu"
	.desc           = "Displays the hash value of the selected file"
	.default_enable = True
	.config.Debug   = False
	.vars.set "Debug", .config.Debug
	.min_version    = "12.0"
		With .AddCommand
			.name     = "ShowHash"
			.method   = "OnShowHash"
			.desc     = ""
			.label    = "ShowHash"
			.hide     = False
			.icon     = "script"
			.template = "TYPE/K[MD5,SHA1,SHA256,SHA512]"
		End With
	End With
End Function

' Called whenever the user modifies the script's configuration
Function OnScriptConfigChange(CData)
	Dim i
	For Each i In CData.changed
		Output i		
	Next
End Function

' Implement the ShowHash command
Function OnShowHash(CmdData)
	Dim Arg, Args, lvItem, File, dlg, msg, ItemIndex, ListView, FilePath, Tab, Files, lvColumn
	Dim MD5, SHA1, SHA256, SHA512

	Set Tab = cmdData.func.sourcetab
    Set dlg = DOpus.Dlg
    dlg.window = Tab
    dlg.template = "Main"
    dlg.detach = True
    dlg.Show
	Set ListView = dlg.Control("lsvList")
	dlg.Control("edtPath").value = Tab.Path
	ListView.columns.AddColumn(DOpus.strings.Get("FilenameCol"))
	
	If CmdData.Func.argsmap.exists("TYPE") Then
		Set Args = Dopus.Create.StringSetI(Split(CmdData.Func.argsmap("TYPE"), ","))
		For Each Arg In Args
			Select Case UCase(Arg)
				Case "MD5"
					MD5 = ListView.columns.AddColumn(Arg) - 1
				Case "SHA1"
					SHA1 = ListView.columns.AddColumn(Arg) - 1
				Case "SHA256"
					SHA256 = ListView.columns.AddColumn(Arg) - 1
				Case "SHA512"
					SHA512 = ListView.columns.AddColumn(Arg) - 1
			End Select
		Next
	Else
		MD5 = ListView.columns.AddColumn("MD5") - 1
		Set Args = Dopus.Create.StringSetI("MD5")
	End If

	For Each lvColumn In ListView.columns
		lvColumn.Resize = True
	Next
	Output Join(Array("MD5:", MD5 , "SHA1:", SHA1 , "SHA256:", SHA256, "SHA512:", SHA512), vbTab)

	If Tab.selected_files.Count > 0 Then
		Set Files = Tab.selected_files
	ElseIf Tab.Files.Count > 0 Then
		Set Files = Tab.Files
	End If

	For Each File In Files
		set lvItem = ListView.GetItemAt(ListView.AddItem(File.name))
		If Args.exists("MD5")    then lvItem.subitems(MD5)    = DOpus.FSUtil.Hash(File.realpath, "md5")
		If Args.exists("SHA1")   then lvItem.subitems(SHA1)   = DOpus.FSUtil.Hash(File.realpath, "SHA1")
		If Args.exists("SHA256") then lvItem.subitems(SHA256) = DOpus.FSUtil.Hash(File.realpath, "SHA256")
		If Args.exists("SHA512") then lvItem.subitems(SHA512) = DOpus.FSUtil.Hash(File.realpath, "SHA512")
	Next
	ListView.columns.AutoSize

	Do
		Set msg = dlg.GetMsg()
		Select Case msg.event
			Case "click"
				select case msg.Control
				case "btnSaveToText"
					set FilePath = DOpus.Dlg.Save("Save","FileHash.txt","#Text Files(*.txt)!*.txt!CSV File(*.csv)!*.csv")
					if FilePath.result then
						Output FilePath
						select case FilePath.ext
						case ".txt"
							WriteTextFile FilePath, Tab.Path & VBCRLF & BuildText(ListView, ""), "UTF-8"
						case ".csv"
							WriteTextFile FilePath, BuildText(ListView, ","), "UTF-8"
						end select
					end if
				case "btnCopy"
						DOpus.SetClip Tab.Path & VBCRLF & BuildText(ListView, vbTab)
				end Select
		End Select 
	Loop While msg

End Function

Function BuildText(ByRef ListView, ByVal Separator)
	Dim lvItem, SubItem, ItemIndex, SubIndex, Result, Lines, Cols
	If Separator ="" Then Separator = " "
	
	For ItemIndex = 0 to ListView.count - 1
		set lvItem = ListView.GetItemAt(ItemIndex)
		Result = Result & lvItem.name
		For Each SubItem in lvItem.SubItems
			Result =  Result & "{DELIMITED}" & SubItem
		Next
		Result = Result & "{CRLF}"
	Next

	Result = Replace(Trim(Result), "{CRLF}", VBCRLF)
	BuildText = Replace(Result, "{DELIMITED}", Separator)
End Function

Function Quotes(ByVal strText)
	Quotes = ChrW(34) + strText + ChrW(34)
End Function

Sub WriteTextFile(FileName, TextContent, Charset)
	Const adSaveCreateNotExist  = 1
	Const adSaveCreateOverWrite = 2
	Const adTypeBinary = 1
	Const adTypeText   = 2
	
	With CreateObject("ADODB.Stream")
		.Type = adTypeText
		.CharSet = Charset
		.Mode = 3
		.Open
		.WriteText TextContent
		.SaveToFile FileName, adSaveCreateOverWrite
		.close
	End With
End Sub

' SCRIPT RESOURCES
==SCRIPT RESOURCES
<resources>
	<resource name="Main" type="dialog">
		<dialog fontface="Microsoft YaHei UI" fontsize="9" height="213" lang="chs" resize="yes" standard_buttons="ok" title="哈希列表" width="258">
			<control fullrow="yes" height="167" name="lsvList" nosortheader="yes" resize="wh" type="listview" viewmode="details" width="246" x="6" y="24" />
			<control height="14" name="btnCopy" resize="y" title="复制到剪贴板" type="button" width="55" x="64" y="195" />
			<control height="14" name="btnSaveToText" resize="y" title="保存到文件" type="button" width="55" x="6" y="195" />
			<control halign="left" height="12" name="edtPath" readonly="yes" resize="w" type="edit" width="246" x="6" y="6" />
		</dialog>
	</resource>
	<resource type="strings">
		<strings lang="chs">
			<string id="FilenameCol">文件名称</string>
		</strings>
	</resource>
</resources>
' SCRIPT RESOURCES