dim a1,a2,b1,b2,c1,c2
a1=2:a2=4
b1=32:b2=67
c1=12:c2=898
msgbox co(a1,a2)
msgbox co(b1,b2)
msgbox co(c1,c2)
function co(t1,t2) '����ʹ��function������һ���µĺ���
	if t1>t2 then
		co=t1 'ͨ��"������=���ʽ"���ַ������ؽ��
	elseif t2>t1 then
		co=t2
	end if
end function