'Count of words and characters in a string

n=Inputbox("Enter a sentence")

str1=Split(n)
msgbox Ubound(str1)+1

for each k in str1
chlen=chlen+Len(k)
next
msgbox "length of chars: " &chlen
