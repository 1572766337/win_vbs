dim h,m,s,t
t=inputbox("请输入秒数：")
h=t\3600
m=t\60 mod 60
s=t mod 60
msgbox h&":"&m&":"&s,,"时分秒"