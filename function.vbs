dim a1,a2,b1,b2,c1,c2
a1=2:a2=4
b1=32:b2=67
c1=12:c2=898
msgbox co(a1,a2)
msgbox co(b1,b2)
msgbox co(c1,c2)
function co(t1,t2) '我们使用function定义了一个新的函数
	if t1>t2 then
		co=t1 '通过"函数名=表达式"这种方法返回结果
	elseif t2>t1 then
		co=t2
	end if
end function