'Renaming a file

set fso=CreateObject("Scripting.FileSystemObject")
source="D:\Personal_Documents\Others\Raaz Reboot (2016) (320 Kbps)"
set oFolder=fso.GetFolder(source)
for each file in oFolder.files
Str=Replace(file.Path," - DownloadMing.SE","")
'msgbox str
fso.MoveFile file.path,Str
next
set fso=nothing

