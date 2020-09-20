a=inputbox("Enter a Number")
if a=1 then
count=0
else for i=2 to a/2 
if a mod i=0 then
count=1
exit for
end if
next
end if

if count=0 then
msgbox "Prime"
else msgbox "Non-Prime"
end if

