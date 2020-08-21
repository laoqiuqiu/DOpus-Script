option explicit

' last_view_mode
' qiuqiu

' This is a script for Directory Opus.
' See https://www.gpsoft.com.au/DScripts/redirect.asp?page=scripts for development information.

' Called by Directory Opus to initialize the script
Function OnInit(initData)
	initData.name = "Remember_last_view_mode"
	initData.version = "1.0"
	initData.copyright = "qiuqiu"
'	initData.url = "https://resource.dopus.com/viewforum.php?f=35"
	initData.desc = "退出时记忆列表窗当前的视图模式"
	initData.default_enable = true
	initData.min_version = "12.0"
	initData.Config.Remember_last_view_mode = True
	If Not DOpus.Vars.Exists("last_view_mode") Then
		DOpus.Create.Command.RunCommand("@Set glob!:last_view_mode=0")
	End If
End Function

' Called when Directory Opus starts up
Function OnOpenLister(OpenListerData)
	If (openListerData.after And Script.Config.Remember_last_view_mode) Then
		Dim Cmd
		Set Cmd = DOpus.Create.Command
		Cmd.SetSourceTab  openListerData.lister.activetab
		Cmd.RunCommand "Set View=" & DOpus.Vars.Get("last_view_mode")
		Set Cmd = Nothing
	Else
		OnOpenLister = True ' Ask to be called again when all the tabs are open.
		Exit Function
	End If	
End Function

' Called when Directory Opus shuts down
Function OnCloseLister(CloseListerData)
	If (DOpus.listers.Count = 1)  Then
		DOpus.Vars.Set "last_view_mode", Replace(DOpus.listers.lastactive.activetab.format.view, "_", "")
	End If
End Function
