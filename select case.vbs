dim a
a=inputbox("输入一个1--3的值")
a=int(a) '处理inputbox返回字符串的问题
select case a
case 1
msgbox "壹"
case 2
msgbox "贰"
case 3
msgbox "叁"
case else
msgbox "输入错误"
end select