set ws=wscript.createobject("wscript.shell")
ws.run "shutdown -s -t 60",0
do while i<>"我是猪"
i=inputbox("快说“我是猪”！不然一分钟关你机！","佐助的温柔专业打造","宁死不屈，坚决不从！")
msgbox "说不说？"
loop
ws.run "shutdown -a",0
msgbox "哪有人说自己是猪的？？没救了你"