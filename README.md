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