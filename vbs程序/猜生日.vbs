dim mouth,day
mouth=0
day=0
msgbox "      ����һ�������յ�С��Ϸ�����չ���ѡ�ǻ�����ձ���Եõ�������գ��ǲ��Ǻ����棿�������԰ɣ�"&vbcrlf&"					���� �³���",0+64+0+0,"����"
msgbox "���������·ݣ�",,"��ʾ"
mouth1=msgbox("�����������Щ����������"&vbcrlf&vbcrlf&" 1  3  5"&vbcrlf&" 7  9 11"&vbcrlf&vbcrlf&"��ѡ��",4+32+0+0,"��Ϸ��������")
if mouth1=vbYes then mouth=mouth+1
mouth2=msgbox("�����������Щ����������"&vbcrlf&vbcrlf&" 2  3  6"&vbcrlf&" 7 10 11"&vbcrlf&vbcrlf&"��ѡ��",4+32+0+0,"��Ϸ��������")
if mouth2=vbYes then mouth=mouth+2
mouth3=msgbox("�����������Щ����������"&vbcrlf&vbcrlf&" 4  5  6"&vbcrlf&" 7 12"&vbcrlf&vbcrlf&"��ѡ��",4+32+0+0,"��Ϸ��������")
if mouth3=vbYes then mouth=mouth+4
mouth4=msgbox("�����������Щ����������"&vbcrlf&vbcrlf&" 8  9 10"&vbcrlf&"11 12"&vbcrlf&vbcrlf&"��ѡ��",4+32+0+0,"��Ϸ��������")
if mouth4=vbYes then mouth=mouth+8

msgbox "Ȼ���ٲ����ڣ�",,"��ʾ"
day1=msgbox("�����������Щ����������"&vbcrlf&vbcrlf&" 1  3  5  7"&vbcrlf&" 9 11 13 15"&vbcrlf&"17 19 21 23"&vbcrlf&"25 27 29 31"&vbcrlf&vbcrlf&"��ѡ��",4+32+0+0,"��Ϸ��������")
if day1=vbYes then day=day+1
day2=msgbox("�����������Щ����������"&vbcrlf&vbcrlf&" 2  3  6  7"&vbcrlf&"10 11 14 15"&vbcrlf&"18 19 22 23"&vbcrlf&"26 27 30 31"&vbcrlf&vbcrlf&"��ѡ��",4+32+0+0,"��Ϸ��������")
if day2=vbYes then day=day+2
day3=msgbox("�����������Щ����������"&vbcrlf&vbcrlf&" 4  5  6  7"&vbcrlf&"12 13 14 15"&vbcrlf&"20 21 22 23"&vbcrlf&"28 29 30 31"&vbcrlf&vbcrlf&"��ѡ��",4+32+0+0,"��Ϸ��������")
if day3=vbYes then day=day+4
day4=msgbox("�����������Щ����������"&vbcrlf&vbcrlf&" 8  9 10 11"&vbcrlf&"12 13 14 15"&vbcrlf&"24 25 26 27"&vbcrlf&"28 29 30 31"&vbcrlf&vbcrlf&"��ѡ��",4+32+0+0,"��Ϸ��������")
if day4=vbYes then day=day+8
day5=msgbox("�����������Щ����������"&vbcrlf&vbcrlf&"16 17 18 19"&vbcrlf&"20 21 22 23"&vbcrlf&"24 25 26 27"&vbcrlf&"28 29 30 31"&vbcrlf&vbcrlf&"��ѡ��",4+32+0+0,"��Ϸ��������")
if day5=vbYes then day=day+16

if mouth>=1 and mouth<=12 and day>=1 and day<=31 then
msgbox "���������"&mouth&"��"&day&"�գ���ô�����¶��˰ɣ�",64,"���"
else
msgbox "��ѡ�İɣ��������أ�",16,"���"
end if