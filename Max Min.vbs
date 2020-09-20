Dim arr(5)
arr(0)=100
arr(1)=70
arr(2)=30
arr(3)=40
arr(4)=50
arr(5)=60
vmax=arr(0)
vmin=arr(0)
l=UBound(arr)
for i=0 to l
if vmin>arr(i) then
vmin=arr(i)

else if vmax<arr(i) then
vmax=arr(i)
end if
end if
Next
msgbox "Min is: "&vmin
msgbox "Max. is: "&vmax

