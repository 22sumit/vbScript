Dim a

a=inputbox("enter string")
rev=""

for i=len(a) to 1 step -1
rev=rev& Mid(a,i,1)
next
msgbox rev
if StrComp(a,rev,1)=0 then
'1 is for texual comparison, 0 is for Binary comparison
msgbox "Palindrome"
else msgbox "non palindrome"
end if

