'get all files with .vbs extension in a folder

Set fso=CreateObject("Scripting.FileSystemObject")
source="E:\Updates\Saathiya 2002"
set ofolder=fso.GetFolder(source)
for each file in ofolder.files
strold=fso.GetFileName(file.path)
strnew=Replace(strold," - (Hsongs.Com)","")
msgbox strnew
next
set fso=nothing
set ofolder=nothing