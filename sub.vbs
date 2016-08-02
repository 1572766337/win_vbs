'有时候我们并不需要返回什么值, 这个时候我们可以使用一种称之为"子程序"的结构. 子程序或称之为过程与函数的差别     就在于:1) 没有返回值, 2) 使用sub关键字定义, 3) 通过Call调用
dim yname
yname=inputbox("请输入你的名字:")
call who(yname)
sub who(cname)
	msgbox "你好" & cname
	msgbox "感谢你阅读我的课程"
	msgbox "这是基础部分的最后一课"
end sub