'on error resume next 
set arg=wscript.arguments '声明外部参数 
if arg.count=0 Then wscript.quit '若无参数则退出脚本 
set fso=createobject("Scripting.FilesystemObject") '声明fso组件 
filename=wscript.arguments(0) 
set readline=fso.opentextfile(filename,1,flase) 
data=readline.readall: 
readline.close '读取文本内容 
if err.number<>0 then wscript.quit '如果发生错误，则退出 
with fso.opentextfile(filename&"_out.vbs",2,true) '将转换好的写到一个新的vbs中 
if err.number<>0 then wscript.quit '如果发生错误，则退出 
.writeline "Execute(Hextostr("""&str2hex(data)&"""))"  '执行解密并执行解密后的代码 
.writeline "Function hextostr(data)" 
.writeline "Hextostr=""Execute""""""""""" 
.writeline "C=""&CHR(&H""" 
.writeline "N= "")""" 
.writeline "Do while Len(data)>1" 
.writeline "if IsNumeric (Left(data,1)) then" 
.writeline "Hextostr=Hextostr&c&Left(data,2)&N" 
.writeline "data = mid(data,3)" 
.writeline "else" 
.writeline "Hextostr=Hextostr&c&Left(data,4)&N" 
.writeline "data=mid(data,5)" 
.writeline "end if" 
.writeline "loop" 
.writeline "end function" 
'把解密函数写进去 
.close '关闭文本 
end with 
set fso=Nothing'注销fso组件 
msgbox "OK" 
'以下是加密函数 
Function str2hex (Byval strHex) 
For i=1 to Len(strHex) 
sHex = sHex & Hex(Asc(mid(strHex,i,1))) 
 next 
 str2Hex = sHex 
end function