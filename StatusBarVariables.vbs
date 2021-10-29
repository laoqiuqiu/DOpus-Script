option explicit

' Status Bar Variables
' qiuqiu

' Called by Directory Opus to initialize the script
Function OnInit(initData)
	initData.name           = "Status Bar Variables"
	initData.version        = "1.0"
	initData.copyright      = "qiuqiu"
	initData.url            = "http://script.dopus.net/"
	initData.desc           = "Status Bar Variables: AvgSize"
	initData.default_enable = True
	initData.min_version    = "12.0"
End Function

' Called when a new tab is opened
Function OnOpenTab(openTabData)
	Dim Tab
	Set Tab = openTabData.Tab
	Tab.vars.Set "AvgSize", CalcAvgSize(Tab)
End Function

' Called when a tab is activated
Function OnActivateTab(activateTabData)
	Dim Tab
	If activateTabData.Result Then
		Set Tab = activateTabData.newtab
		Tab.vars.Set "AvgSize", CalcAvgSize(Tab)
	End If
End Function

' Called after a new folder is read in a tab
Function OnAfterFolderChange(afterFolderChangeData)
	Dim Tab
	If afterFolderChangeData.Result Then
		Set Tab = afterFolderChangeData.Tab
		Tab.vars.Set "AvgSize", CalcAvgSize(Tab)
	End If
End Function

Function CalcAvgSize(ByRef Tab)
	Dim Files_Size
	Set Files_Size = DOpus.FSUtil.NewFileSize(Tab.Stats.FileBytes)
	If Files_Size.CY > 0 Then
		Files_Size.Div Tab.Stats.Files
		CalcAvgSize = Files_Size.Fmt
	End If
End Function
