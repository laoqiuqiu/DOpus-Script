# DOpus-Script
Directory Opus 脚本

## ShowFilesHash.vbs
显示文件的哈希值的命令
```
ShowHash TYPE=MD5,SHA1 
```
## timeago.vbs
* 两个自定义列，显示文件或文件的创建和修改时间，例如：1 小时前

## StatusBarVariables.vbs
状态栏变量，`{var:tab:AvgSize}` 计算当前标签中文件平均大小(不包含文件夹)

## CopyContent.vbs
* 将文件内容发送到剪贴板
`CopyContent c:\test.txt`
`CopyContent c:\temp PATTERN *.(txt|vbs)`

仅将选中文件中的扩展名为 txt 的文件内容发送到剪贴板
`CopyContent PATTERN *.TXT`

## DOpus.ObjectDefs.vbs
适用于 <a href="https://github.com/Serpen/VBS-VSCode" target="_blank">VBScript Extension for Visual Studio Code</a> Directory Opus 脚本对象扩展

**在扩展设置中添加**
```
{ // settings.json
    "vbs.includes": ["x:\\xxxx\\DOpus.ObjectDefs.vbs"]
}
```
