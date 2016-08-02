dim i,j
for i=1 to 9
	for j=1 to 9
		str=str & i * j & " " '&是和并字符串的符号
	next '每个next对应一个for
	str=str & vbCrlf 'vbCrlf相当于键盘上的回车键,因为你不能在键盘上输入,所以系统定义了一个默认的常量
next
msgbox str