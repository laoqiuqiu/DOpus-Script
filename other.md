# 自定义功能集合
我在使用 Directory Opus 过程中用到的一些东东

### BandZIP 右键菜单
```
FileType CONTEXTMENU={5B69A6B4-393B-459C-8EBB-214237A9E7AC} CONTEXTFORCE 
```
### 在资源管理器中打开
```
@hidenosel:maxfiles=1,maxdirs=1
"/windows\explorer.exe" /Select, /e, {s}
```
### 删除 GPS 信息
```
SetAttr META gpsaltitude gpslatitude gpslongitude
```
### DOpus 13 程序分组提示信息
```
<b><#32CD32>{name}</#></b>{thumbnail}
 
<b>文件大小：</b>\	{sizeauto}({size} Bytes)
<b>出品公司：</b>\	{companyname}
<b>产品版本：</b>\	{prodversion}
<b>版权信息：</b>\	{copyright}
<b>文件描述：</b>\	{moddesc}
<b>文件版本：</b>\	{modversion}
<b>创建时间：</b>\	{created}
<b>修改时间：</b>\	{modified}
\	
<b>程序架构：</b>\	{=Return(RegEx(Mid(desc, (Len(userdesc) != 0 ? Len(userdesc) + 3: 0)), "(^\w*), (.*)", "\1"));=}{!=return(Match(desc, "(\.net)", "rp"))=} (.NET){!}
<b>数字签名：</b>\	{=return(RegEx(desc, "(" + LanguageStr(5650) + "|" + LanguageStr(5651) + ")(.*)", "\1"));=}
```

### DOpus 13 求值器列
```
<?xml version="1.0"?>
<evalcolumn align="0" attrrefresh="no" autorefresh="no" category="prog" customgrouping="no" foldertype="shell" keyword="Is.NET" maxstars="5" namerefresh="no" reversesort="no" title=".NET" type="0">return(UCase(RegEx(desc, &quot;(\w\d*?\W\s)(\.NET)(\W\s)&quot;, &quot;&quot;) == &quot;&quot;));</evalcolumn>

<?xml version="1.0"?>
<evalcolumn align="0" attrrefresh="no" autorefresh="no" category="prog" customgrouping="no" foldertype="shell" header="平台" keyword="Platform" maxstars="5" namerefresh="no" reversesort="no" title="平台" type="0">return(RegEx(desc, &quot;(^\w+\d*)(.*)&quot;, &quot;\1&quot;));</evalcolumn>

<?xml version="1.0"?>
<evalcolumn align="0" attrrefresh="no" autorefresh="no" category="prog" customgrouping="no" foldertype="shell" header="数字签名" keyword="Signature" maxstars="5" namerefresh="no" reversesort="no" title="数字签名" type="0">return(RegEx(desc, &quot;\.NET\W\s(\w+)&quot;, &quot;\1&quot;, &quot;(^\w+\d*)\W\s(\w+)(\W\s)*&quot;, &quot;\2&quot;));</evalcolumn>

<?xml version="1.0"?>
<evalcolumn align="1" attrrefresh="yes" autorefresh="yes" category="date" customgrouping="no" foldertype="all" graphcol="#ff8000" header="创建于" keyword="CreateAt" maxstars="5" namerefresh="yes" reversesort="no" title="创建于" type="0">	seconds = DateDiff(&quot;s&quot;, created, Now());
	Suffix = (seconds &gt;= 0) ? &quot; Ago&quot; : &quot; Later&quot;;
	
	s = seconds Mod 60;
	n = seconds / 60 Mod 60;
	h = seconds / 60 / 60 Mod 24;
	d = seconds / 60 / 60 / 24 mod 30;
	m = seconds / 60 / 60 / 24 / 30 mod 12;
	y = seconds / 60 / 60 / 24 / 30 / 12;
	
	ss = (d &gt; 0 || s == 0) ? &quot;&quot; : s + &quot;s&quot;;
	nn = (m == 0) ? &quot;&quot; : m + &quot;m &quot;;
	hh = (h == 0) ? &quot;&quot; : h + &quot;h &quot;;
	dd = (d == 0) ? &quot;&quot; : d + &quot;d &quot;;
	mm = (m == 0) ? &quot;&quot; : m + &quot;m &quot;;
	yy = (y == 0) ? &quot;&quot; : y + &quot;y &quot;;
	
	return Trim(yy + dd + hh + nn + ss + Suffix);
</evalcolumn>

<?xml version="1.0"?>
<evalcolumn align="1" attrrefresh="yes" autorefresh="yes" category="date" customgrouping="no" foldertype="all" graphcol="#ff8000" header="修改于" keyword="ModifyAt" maxstars="5" namerefresh="yes" reversesort="no" title="修改于" type="0">	seconds = DateDiff(&quot;s&quot;, modified, Now());
	Suffix = (seconds &gt;= 0) ? &quot; Ago&quot; : &quot; Later&quot;;
	
	s = seconds Mod 60;
	n = seconds / 60 Mod 60;
	h = seconds / 60 / 60 Mod 24;
	d = seconds / 60 / 60 / 24 mod 30;
	m = seconds / 60 / 60 / 24 / 30 mod 12;
	y = seconds / 60 / 60 / 24 / 30 / 12;
	
	ss = (d &gt; 0 || s == 0) ? &quot;&quot; : s + &quot;s&quot;;
	nn = (m == 0) ? &quot;&quot; : m + &quot;m &quot;;
	hh = (h == 0) ? &quot;&quot; : h + &quot;h &quot;;
	dd = (d == 0) ? &quot;&quot; : d + &quot;d &quot;;
	mm = (m == 0) ? &quot;&quot; : m + &quot;m &quot;;
	yy = (y == 0) ? &quot;&quot; : y + &quot;y &quot;;
	
	return Trim(yy + dd + hh + nn + ss + Suffix);
	
</evalcolumn>
```
