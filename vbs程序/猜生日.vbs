dim mouth,day
mouth=0
day=0
msgbox "      这是一个猜生日的小游戏，按照规则选是或否，最终便可以得到你的生日，是不是很神奇？快来试试吧！"&vbcrlf&"					—— 陕晨阳",0+64+0+0,"介绍"
msgbox "首先来猜月份！",,"提示"
mouth1=msgbox("你的生日在这些数字里面吗？"&vbcrlf&vbcrlf&" 1  3  5"&vbcrlf&" 7  9 11"&vbcrlf&vbcrlf&"请选择：",4+32+0+0,"游戏：猜生日")
if mouth1=vbYes then mouth=mouth+1
mouth2=msgbox("你的生日在这些数字里面吗？"&vbcrlf&vbcrlf&" 2  3  6"&vbcrlf&" 7 10 11"&vbcrlf&vbcrlf&"请选择：",4+32+0+0,"游戏：猜生日")
if mouth2=vbYes then mouth=mouth+2
mouth3=msgbox("你的生日在这些数字里面吗？"&vbcrlf&vbcrlf&" 4  5  6"&vbcrlf&" 7 12"&vbcrlf&vbcrlf&"请选择：",4+32+0+0,"游戏：猜生日")
if mouth3=vbYes then mouth=mouth+4
mouth4=msgbox("你的生日在这些数字里面吗？"&vbcrlf&vbcrlf&" 8  9 10"&vbcrlf&"11 12"&vbcrlf&vbcrlf&"请选择：",4+32+0+0,"游戏：猜生日")
if mouth4=vbYes then mouth=mouth+8

msgbox "然后再猜日期！",,"提示"
day1=msgbox("你的生日在这些数字里面吗？"&vbcrlf&vbcrlf&" 1  3  5  7"&vbcrlf&" 9 11 13 15"&vbcrlf&"17 19 21 23"&vbcrlf&"25 27 29 31"&vbcrlf&vbcrlf&"请选择：",4+32+0+0,"游戏：猜生日")
if day1=vbYes then day=day+1
day2=msgbox("你的生日在这些数字里面吗？"&vbcrlf&vbcrlf&" 2  3  6  7"&vbcrlf&"10 11 14 15"&vbcrlf&"18 19 22 23"&vbcrlf&"26 27 30 31"&vbcrlf&vbcrlf&"请选择：",4+32+0+0,"游戏：猜生日")
if day2=vbYes then day=day+2
day3=msgbox("你的生日在这些数字里面吗？"&vbcrlf&vbcrlf&" 4  5  6  7"&vbcrlf&"12 13 14 15"&vbcrlf&"20 21 22 23"&vbcrlf&"28 29 30 31"&vbcrlf&vbcrlf&"请选择：",4+32+0+0,"游戏：猜生日")
if day3=vbYes then day=day+4
day4=msgbox("你的生日在这些数字里面吗？"&vbcrlf&vbcrlf&" 8  9 10 11"&vbcrlf&"12 13 14 15"&vbcrlf&"24 25 26 27"&vbcrlf&"28 29 30 31"&vbcrlf&vbcrlf&"请选择：",4+32+0+0,"游戏：猜生日")
if day4=vbYes then day=day+8
day5=msgbox("你的生日在这些数字里面吗？"&vbcrlf&vbcrlf&"16 17 18 19"&vbcrlf&"20 21 22 23"&vbcrlf&"24 25 26 27"&vbcrlf&"28 29 30 31"&vbcrlf&vbcrlf&"请选择：",4+32+0+0,"游戏：猜生日")
if day5=vbYes then day=day+16

if mouth>=1 and mouth<=12 and day>=1 and day<=31 then
msgbox "你的生日是"&mouth&"月"&day&"日，怎么样，猜对了吧？",64,"结果"
else
msgbox "乱选的吧，逗我玩呢？",16,"结果"
end if