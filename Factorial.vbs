Function Factorial(x)
f=1
for i=1 to x
f=f*i
next
Factorial=f
End Function

a=Inputbox ("Enter a number")
msgbox ("Factorial of " & a &" is " &Factorial (a))