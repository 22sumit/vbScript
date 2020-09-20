Dim arr()
ReDim arr(5)
arr(0)=10
arr(1)=20
arr(2)=30
arr(3)=40
arr(4)=50
arr(5)=60
l=UBound(arr)
for i=0 to l
msgbox arr(i)
Next

msgbox "ReDim Using Preserve"
Redim Preserve arr(6)
arr(6)=70
for i=0 to UBound(arr)
msgbox arr(i)
Next
