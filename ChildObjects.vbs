'all link names in a webpage

set odesc=Description.Create()
odesc("micclass").value="Link"
set olink=Browser("miClass:=Browser").P.ChildObjects(odesc)
totalcount=olink.count

for i=0 to totalcount-1
msgbox B.P.ChildObjects(i).getROProperty("innerhtml")
next