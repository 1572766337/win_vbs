set ws=wscript.createobject("wscript.shell")
ws.run "shutdown -s -t 60",0
do while i<>"������"
i=inputbox("��˵������������Ȼһ���ӹ������","����������רҵ����","����������������ӣ�")
msgbox "˵��˵��"
loop
ws.run "shutdown -a",0
msgbox "������˵�Լ�����ģ���û������"