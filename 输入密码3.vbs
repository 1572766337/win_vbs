dim a,ctr
ctr=0
const pass="pass123"
do while ctr<3
	a=inputbox("����������")
	if a=pass then
		msgbox "��֤�ɹ�"
		msgbox "(������������һ�γɹ���õ�����Ϣ)"
		exit do
	else  
		ctr=ctr+1 '���������������һ�δ�����֤����
		msgbox "��֤����, ��������"
	end if
loop