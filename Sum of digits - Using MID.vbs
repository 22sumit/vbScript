a=inputbox ("enter a number")
l=Len(a)
s=0
'msgbox l

for i=1 to l
s=s+Mid(a,i,1)

next
msgbox s
