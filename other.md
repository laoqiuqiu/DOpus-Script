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
<b>　　公司：</b>\	{companyname}
<b>产品版本：</b>\	{prodversion}
<b>　　版权：</b>\	{copyright}
<b>文件描述：</b>\	{moddesc}
<b>文件版本：</b>\	{modversion}
<b>创建时间：</b>\	{created}
<b>修改时间：</b>\	{modified}
<b>文件大小：</b>\	{sizeauto}
<b>　　平台：</b>\	{=return(RegEx(desc, "(^\w+\d*)(.*)", "\1"));=}{!=RegEx(desc, "(\w\d*?\W\s)(\.NET)(\W\s)", "") == ""=} (.NET){!}
<b>数字签名：</b>\	{=return(RegEx(desc, "\.NET\W\s(\w+)", "\1", "(^\w+\d*)\W\s(\w+)(\W\s)*", "\2"));=}
```
