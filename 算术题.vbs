Dim a,b,c,d
msgbox "a*2+b*2=��",,"����"
a=inputbox("a��:","����뾶")
b=Inputbox("b��:","����뾶")
d=Inputbox("a*2+b*2=:","�����")
d=int(d)
'����������ȡ����d��ֵ, �������, �ٷŻ�"d"������������d��c���Ͳ�ƥ�䣬��Զ����
c=a*2+b*2
if d=c then
msgbox "����� ��ô�����"
else
msgbox "��ͷ �Լ����⻹���ᣡ"
end if