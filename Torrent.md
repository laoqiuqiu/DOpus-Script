```vbscript
Const s1 = "lli1321372064eel32:udp://tracker.ccc.de:80/announceel38:udp://tracker.publicbt.com:80/announceel44:udp://tracker.openbittorrent.com:80/announceel31:udp://9.rarbg.com:2710/announceel30:udp://11.rarbg.com:80/announceel39:http://bt.careland.com.cn:6969/announceel31:http://cpleft.com:2710/announceel28:http://10.rarbg.com/announceel38:http://exodus.desync.com:6969/announceel24:http://pow7.com/announceee"
Const s2 = "d13:creation datei1321372064e8:encoding5:UTF-8e"
Const F1 = "D:\Temp\Grotesque+Tactics+2+Dungeons+and+Donuts-FLT.torrent"

Set Stream = New BinaryStream
Stream.LoadFromFile(F1)
'Stream.LoadFromString(s2)

Dim T, C, i
C = Stream.ReadChar
Set T = decode(Stream, C)

wsh.echo "main dict count: " & T.count
wsh.echo "main dict keys : " & Join(T.keys, ",")
wsh.echo "info dict keys : " & Join(T("info").keys, ",")
wsh.echo "files dict keys: " & Join(T("info")("files")(0).Keys)
wsh.echo "file count     : " & T("info")("files").count
wsh.echo "patch count    : " & T("info")("files")(0)("path").count

wsh.echo vbNewLine & "announce-list"
wsh.echo "list count      : " & T("announce-list").count
wsh.echo "list item count : " & T("announce-list")(0).count
For i = 0 To T("announce-list").count - 1
  wsh.echo "list item       : " & T("announce-list")(i)(0)
next
```

output
```
main dict count: 7
main dict keys : announce,announce-list,created by,creation date,encoding,info,nodes
info dict keys : files,name,name.utf-8,piece length,pieces,publisher,publisher-url,publisher-url.utf-8,publisher.utf-8
files dict keys: ed2k filehash length path path.utf-8
file count     : 29
patch count    : 1

announce-list
list count      : 347
list item count : 1
list item       : http://0d.kebhana.mx:443/announce
list item       : http://104.238.198.186:8000/announce
list item       : http://104.28.1.30:8080/announce
list item       : http://104.28.16.69/announce
list item       : http://107.150.14.110:6969/announce
list item       : http://109.121.134.121:1337/announce
list item       : http://114.55.113.60:6969/announce
......
```