dim a,ctr
ctr=0 '���ü�����
const pass="pass123" '������Ǹ���������, ��θĵ�ǿһ��
do
	a=inputbox("����������")
	if a=pass then
		msgbox "��֤�ɹ�"
		exit do
	elseif ctr=3 then
			msgbox "�Ѿ��ﵽ��֤����, ��֤����ر�"
			exit do
		else
			ctr=ctr+1
			msgbox "��֤����, ��������"
		end if
	'end if
loop