set fso=CreateObject("Scripting.FileSystemObject")
set qfile=fso.GetFolder("D:\Padhai\QTP & UFT Docs\VBScript Programs")
for each file in qfile.Files
if LCase(fso.GetExtensionName(file))="vbs" then
count=count+1
end if
next
msgbox count
set fso=nothing
set qfile=nothing