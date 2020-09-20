Function fibonacci(n)
Dim arr
Redim arr(n)
arr(0)=1
arr(1)=1
for i=2 to n
arr(i)=arr(i-1)+arr(i-2)
Next
fibonacci=arr
End Function

x=Inputbox ("Enter a number")
myarr=fibonacci(x)

for j=0 to Ubound(myarr)-1
msgbox myarr(j)
next
