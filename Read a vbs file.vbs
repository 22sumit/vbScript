set fso=createObject("Scripting.FileSystemObject")
Set qfile=fso.OpenTextFile("D:\Sumit\Personal\fb.txt",8,True)
qfile.WriteLine "All the Best"
Set qfile=fso.OpenTextFile("D:\Sumit\Personal\fb.txt",1,True)
Do while qfile.AtEndofStream<>True
msgbox qfile.Read(3)
Loop
qfile.close
set qfile=nothing
set fso=nothing 