dim i,j
for i=1 to 9
	for j=1 to 9
		str=str & i * j & " " '&�ǺͲ��ַ����ķ���
	next 'ÿ��next��Ӧһ��for
	str=str & vbCrlf 'vbCrlf�൱�ڼ����ϵĻس���,��Ϊ�㲻���ڼ���������,����ϵͳ������һ��Ĭ�ϵĳ���
next
msgbox str