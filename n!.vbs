on error resume next
dim a,b,c,s
b=1
c=2
a=inputbox("���������������","����׳�")
a=int(a)
do while c<=a
b=b*c
c=c+1
loop
msgbox a&"!="&b,,"�׳�"
