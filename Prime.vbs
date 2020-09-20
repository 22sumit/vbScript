count=False
x=InputBox ("Enter the number")

if x=1 then
count=False

else if x>=2 then

for i=2 to x/2

if (x mod i=0) then

count=True

end if

Next

end if

end if

If count=False Then
msgbox "the entered number is prime: "&a
else
msgbox "the entered no. is not prime: "&a
End If