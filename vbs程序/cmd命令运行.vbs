command=inputbox("����Ҫ���е�cmd���")
set ws=createobject("wscript.shell")
ws.run "cmd /c"&command