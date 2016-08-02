command=inputbox("输入要运行的cmd命令：")
set ws=createobject("wscript.shell")
ws.run "cmd /c"&command