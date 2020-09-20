'Read 4th line from a text file

set fso=CreateObject("Scripting.FileSystemObject")
set qfile=fso.OpentextFile("D:\testvbs.txt",1,True)
for i=1 to 3
qfile.readline
next
msgbox qfile.readline
set fso=nothing
set qfile=nothing
