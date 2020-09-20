'Renaming a file

set fso=CreateObject("Scripting.FileSystemObject")
source="D:\Personal_Documents\Padhai\VBScript Programs"
set oFolder=fso.GetFolder(source)
for each file in oFolder.files
Str=Replace(file.Path,fso.getExtensionName(file.Name),"txt")
'msgbox str
fso.MoveFile file.path,Str
next
set fso=nothing

