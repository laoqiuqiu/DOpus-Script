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
