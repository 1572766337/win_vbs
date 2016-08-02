on error resume next
set objshell=createobject("Wscript.shell")
objshell.Run "notpad",,true
objShell.Run "calc"