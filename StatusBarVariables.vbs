option explicit

' Status Bar Variables
' qiuqiu

' This is a script for Directory Opus.
' See https://www.gpsoft.com.au/DScripts/redirect.asp?page=scripts for development information.



' Called by Directory Opus to initialize the script
Function OnInit(initData)
	initData.name = "Status Bar Variables"
	initData.version = "1.0"
	initData.copyright = "qiuqiu"
'	initData.url = "https://resource.dopus.com/viewforum.php?f=35"
	initData.desc = "Status Bar Variables: AvgSize"
	initData.default_enable = true
	initData.min_version = "12.0"
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
	Dim MaxSize, TotalSize, TabItem
	Set MaxSize = DOpus.FSUtil.NewFileSize
	Set TotalSize = DOpus.FSUtil.NewFileSize
	If Tab.Files.Count > 0 Then
		For Each TabItem In Tab.Files
			If TabItem.Size > MaxSize Then MaxSize = TabItem.Size
			If (not TabItem.is_junction) And (Not TabItem.is_reparse) And (Not TabItem.is_symlink) Then TotalSize.Add TabItem.Size
		Next
		TotalSize.Div Tab.Files.Count
		CalcAvgSize = TotalSize.Fmt
	End If
End Function
