option explicit

' Document pages count
' (c) 2020 qiuqiu

' Called by Directory Opus to initialize the script
Function OnInit(initData)
	Dim props
	Set props = DOpus.FSUtil.GetShellPropertyList("System.Document.PageCount", "r")

	with initData
		.name           = "Document pages count"
		.version        = "1.0"
		.copyright      = "qiuqiu"
		.desc           = DOpus.Strings.Get("ScriptDesc")
		.url            = "http://script.dopus.net/"
		.default_enable = True
		.min_version    = "12.0.8"
 
		with .AddColumn
			.name     = props(0).raw_name
			.method      = "On_DocPages"
			.label       = DOpus.Strings.Get("ColumnLabel")
			.justify     = props(0).justify
			.type        = props(0).Type
			.autogroup   = True
			.autorefresh = True
			.userdata = props(0).pkey
		end with

	end with

End Function


' Implement the DocPages column
Function On_DocPages(scriptColData)
	scriptColData.value = scriptColData.item.shellprop(scriptColData.userdata)
End Function

==SCRIPT RESOURCES
<resources>
    <resource type = "Strings">
        <Strings lang = "english">
            <string id = "ScriptDesc"  text = "Add the document page number column in the shell extension." />
            <string id = "ColumnLabel" text = "Pages" />
		</Strings>
		<Strings lang = "chs">
            <string id = "ScriptDesc"  text = "添加外壳扩展中的文档页数列." />
            <string id = "ColumnLabel" text = "页数" />
        </Strings>
    </resource>
</resources>
