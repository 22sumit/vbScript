'get the count of files having extension xlxs in a folder

Set fso=CreateObject("Scripting.FileSystemObject")
source="D:\Florida\QEVAL"
set folder=fso.GetFolder(source)
count=0
for each ofile in folder.files
if Lcase(fso.GetExtensionName(ofile.Name))="xlsx" then
count=count+1
end if
next
msgbox count
