@echo off
start mshta VBScript:createobject("sapi.spvoice").speak("导入成功")(close)
::mshta VBScript:msgbox("导入成功！")(close)