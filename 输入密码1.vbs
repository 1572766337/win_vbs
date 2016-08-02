dim a '注意:常量不需要在dim里面声明,否则会引发错误
const pass="123456" '这是一个字符串 请用""包裹起来. 设定密码为常量, 不可变更
do
	a=inputbox("请输入密码")
	if a=pass then
		msgbox "密码校验成功"
		exit do
	end if
loop