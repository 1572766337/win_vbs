on error resume next
dim a,b,c
b=1
c=2
a=inputbox("请输入运算次数：","计算1+2+…+n")
a=int(a)
do while c<=a
b=b+c
c=c+1
loop
msgbox "1+2+…+"&a&"="&b,,"结果"
