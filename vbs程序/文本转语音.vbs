msgbox "本程序可以朗读输入的文本，输入“退出”即可退出！",,"程序介绍"
do while a<>"退出"
set ws=createobject("sapi.spvoice")
a=inputbox("输入要朗读的单词：")
ws.speak a
loop
