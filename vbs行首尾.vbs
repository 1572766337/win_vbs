'on error resume next 
set arg=wscript.arguments '声明外部参数 
if arg.count=0 Then
msgbox "请拖入要修改的文件！",64+4096,"提示"
wscript.quit '若无参数则退出脚本
end if
homeline=inputbox("输入要添加在行首的词：")
endline=inputbox("输入要添加在末尾的词：")
set fso=createobject("Scripting.FilesystemObject") '声明fso组件 
filename=wscript.arguments(0)
set readf=fso.opentextfile(filename,1,flase)
do until readf.AtEndOfStream
data=readf.readline
if err.number<>0 then wscript.quit '如果发生错误，则退出 
with fso.opentextfile(filename&"_out.txt",8,true) '将转换好的写到一个新的vbs中(8追加)
if err.number<>0 then wscript.quit '如果发生错误，则退出 
.write homeline&data&endline&vbcrlf
end with
loop
set fso=Nothing'注销fso组件 
msgbox "OK，转换成功！" 