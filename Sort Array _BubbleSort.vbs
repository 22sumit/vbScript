'bubble sort

dim arr(4)
arr(0)=10
arr(1)=200
arr(2)=40
arr(3)=80
arr(4)=50
for j=0 to ubound(arr)
for i=0 to ubound(arr)-1
if arr(i)<arr(i+1) then
temp=arr(i)
arr(i)=arr(i+1)
arr(i+1)=temp
end if
next
next
for i=0 to ubound(arr)
msgbox arr(i)
next