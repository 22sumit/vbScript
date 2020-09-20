'Hide a file in windows

set fso=CreateObject("Scripting.FileSystemObject")
set qfile=fso.getFile("D:\testvbscopy.txt")
qfile.Attributes = 2 'this line hides the file
set fso=nothing
set qfile=nothing