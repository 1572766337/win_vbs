'on error resume next 
set arg=wscript.arguments '�����ⲿ���� 
if arg.count=0 Then
msgbox "������Ҫ�޸ĵ��ļ���",64+4096,"��ʾ"
wscript.quit '���޲������˳��ű�
end if
homeline="echo "
endline=inputbox("ע�⣺"&vbcrlf&vbcrlf&"    ����Ҫ��ӵ��ļ�·�������֣�","��β")
set fso=createobject("Scripting.FilesystemObject") '����fso��� 
filename=wscript.arguments(0)
set readf=fso.opentextfile(filename,1,flase)
do until readf.AtEndOfStream
data=readf.readline
if err.number<>0 then wscript.quit '��������������˳�
with fso.opentextfile(filename&"_out.bat",8,true) '��ת���õ�д��һ���µ�vbs��(8׷��)
if err.number<>0 then wscript.quit '��������������˳�
.write homeline&data&">>"&endline&vbcrlf
end with
loop
set fso=Nothing'ע��fso���
msgbox "OK��ת���ɹ���"