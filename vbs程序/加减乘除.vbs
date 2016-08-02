randomize
msgbox "这是一个计算题程序，可以做加、减、乘、除、乘方运算",,"介绍"
do
calc=inputbox("1.加法"&vbcrlf&"2.减法"&vbcrlf&"3.乘法"&vbcrlf&"4.除法"&vbcrlf&"5.乘方"&vbcrlf&"6.退出"&vbcrlf&"请选择所要进行的运算序号：","菜单")
range=inputbox("1.1到9"&vbcrlf&"2.10到99"&vbcrlf&"3.100到999"&vbcrlf&"请选择所要计算的数值范围：","范围")
n1=int(9*rnd)+1
n2=int(9*rnd)+1
n3=int(89*rnd)+10
n4=int(89*rnd)+10
n5=int(899*rnd)+100
n6=int(899*rnd)+100
if calc="1" and range="1" then
if int(inputbox(n1&"+"&n2&"=？"))=n1+n2 then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&n1+n2
end if
elseif calc="2" and range="1" then
if int(inputbox(n1&"-"&n2&"=？"))=n1-n2 then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&n1-n2
end if
elseif calc="3" and range="1" then
if int(inputbox(n1&"x"&n2&"=？"))=n1*n2 then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&n1*n2
end if
elseif calc="4" and range="1" then
if int(inputbox(n1&"/"&n2&"=？"))=round(n1/n2,2) then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&round(n1/n2,2)
end if
elseif calc="5" and range="1" then
if int(inputbox(n1&"x"&n1&"=？"))=n1*n1 then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&n1*n1
end if
elseif calc="1" and range="2" then
if int(inputbox(n3&"+"&n4&"=？"))=n3+n4 then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&n3+n4
end if
elseif calc="2" and range="2" then
if int(inputbox(n3&"-"&n4&"=？"))=n3-n4 then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&n3-n4
end if
elseif calc="3" and range="2" then
if int(inputbox(n3&"x"&n4&"=？"))=n3*n4 then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&n3*n4
end if
elseif calc="4" and range="2" then
if int(inputbox(n3&"/"&n4&"=？"))=round(n3/n4,2) then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&round(n3/n4,2)
end if
elseif calc="5" and range="2" then
if int(inputbox(n3&"x"&n3&"=？"))=n3*n3 then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&n3*n3
end if
elseif calc="1" and range="3" then
if int(inputbox(n5&"+"&n6&"=？"))=n5+n6 then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&n5+n6
end if
elseif calc="2" and range="3" then
if int(inputbox(n5&"-"&n6&"=？"))=n5-n6 then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&n5-n6
end if
elseif calc="3" and range="3" then
if int(inputbox(n5&"x"&n6&"=？"))=n5*n6 then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&n5*n6
end if
elseif calc="4" and range="3" then
if int(inputbox(n5&"/"&n6&"=？"))=round(n5/n6,2) then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&round(n5/n6,2)
end if
elseif calc="5" and range="3" then
if int(inputbox(n5&"x"&n5&"=？"))=n5*n5 then
msgbox "恭喜，你做对了！"
else
msgbox "笨蛋，答错了！"&"正确答案是："&n5*n5
end if
elseif calc="6" then exit do
end if
loop