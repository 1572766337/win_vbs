msgbox "����������ʶ�������ı������롰�˳��������˳���",,"�������"
do while a<>"�˳�"
set ws=createobject("sapi.spvoice")
a=inputbox("����Ҫ�ʶ��ĵ��ʣ�")
ws.speak a
loop
