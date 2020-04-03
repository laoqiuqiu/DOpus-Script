option explicit

' ProtectLockedTabs
' qiuqiu

' This is a script for Directory Opus.
' See https://www.gpsoft.com.au/DScripts/redirect.asp?page=scripts for development information.



' Called by Directory Opus to initialize the script
Function OnInit(initData)
	With initData
		.name           = "ProtectLockedTabs"
		.version        = "1.0"
		.copyright      = "qiuqiu"
		.url            = "http://script.dopus.net/"
		.desc           = "Prevent the locked tab from closing,"
		.default_enable = true
		.min_version    = "12.0"
	End With
End Function

' Called when a tab is closed
Function OnCloseTab(TabData)
	If TabData.Tab.Lock <> "off" Then OnCloseTab = True
End Function

