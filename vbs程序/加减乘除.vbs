randomize
msgbox "����һ����������򣬿������ӡ������ˡ������˷�����",,"����"
do
calc=inputbox("1.�ӷ�"&vbcrlf&"2.����"&vbcrlf&"3.�˷�"&vbcrlf&"4.����"&vbcrlf&"5.�˷�"&vbcrlf&"6.�˳�"&vbcrlf&"��ѡ����Ҫ���е�������ţ�","�˵�")
range=inputbox("1.1��9"&vbcrlf&"2.10��99"&vbcrlf&"3.100��999"&vbcrlf&"��ѡ����Ҫ�������ֵ��Χ��","��Χ")
n1=int(9*rnd)+1
n2=int(9*rnd)+1
n3=int(89*rnd)+10
n4=int(89*rnd)+10
n5=int(899*rnd)+100
n6=int(899*rnd)+100
if calc="1" and range="1" then
if int(inputbox(n1&"+"&n2&"=��"))=n1+n2 then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&n1+n2
end if
elseif calc="2" and range="1" then
if int(inputbox(n1&"-"&n2&"=��"))=n1-n2 then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&n1-n2
end if
elseif calc="3" and range="1" then
if int(inputbox(n1&"x"&n2&"=��"))=n1*n2 then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&n1*n2
end if
elseif calc="4" and range="1" then
if int(inputbox(n1&"/"&n2&"=��"))=round(n1/n2,2) then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&round(n1/n2,2)
end if
elseif calc="5" and range="1" then
if int(inputbox(n1&"x"&n1&"=��"))=n1*n1 then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&n1*n1
end if
elseif calc="1" and range="2" then
if int(inputbox(n3&"+"&n4&"=��"))=n3+n4 then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&n3+n4
end if
elseif calc="2" and range="2" then
if int(inputbox(n3&"-"&n4&"=��"))=n3-n4 then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&n3-n4
end if
elseif calc="3" and range="2" then
if int(inputbox(n3&"x"&n4&"=��"))=n3*n4 then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&n3*n4
end if
elseif calc="4" and range="2" then
if int(inputbox(n3&"/"&n4&"=��"))=round(n3/n4,2) then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&round(n3/n4,2)
end if
elseif calc="5" and range="2" then
if int(inputbox(n3&"x"&n3&"=��"))=n3*n3 then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&n3*n3
end if
elseif calc="1" and range="3" then
if int(inputbox(n5&"+"&n6&"=��"))=n5+n6 then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&n5+n6
end if
elseif calc="2" and range="3" then
if int(inputbox(n5&"-"&n6&"=��"))=n5-n6 then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&n5-n6
end if
elseif calc="3" and range="3" then
if int(inputbox(n5&"x"&n6&"=��"))=n5*n6 then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&n5*n6
end if
elseif calc="4" and range="3" then
if int(inputbox(n5&"/"&n6&"=��"))=round(n5/n6,2) then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&round(n5/n6,2)
end if
elseif calc="5" and range="3" then
if int(inputbox(n5&"x"&n5&"=��"))=n5*n5 then
msgbox "��ϲ���������ˣ�"
else
msgbox "����������ˣ�"&"��ȷ���ǣ�"&n5*n5
end if
elseif calc="6" then exit do
end if
loop