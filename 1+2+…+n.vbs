on error resume next
dim a,b,c
b=1
c=2
a=inputbox("���������������","����1+2+��+n")
a=int(a)
do while c<=a
b=b+c
c=c+1
loop
msgbox "1+2+��+"&a&"="&b,,"���"
