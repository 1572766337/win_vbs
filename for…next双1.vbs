dim i,j
for i=1 to 9
	for j=1 to 9
		str=str & i * j & " " '&是和并字符串的符号
	next '每个next对应一个for
next
msgbox str
