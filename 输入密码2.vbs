dim a,ctr
ctr=0 '设置计数器
const pass="pass123" '上面的那个是弱密码, 这次改的强一点
do
	a=inputbox("请输入密码")
	if a=pass then
		msgbox "认证成功"
		exit do
	elseif ctr=3 then
			msgbox "已经达到认证上限, 认证程序关闭"
			exit do
		else
			ctr=ctr+1
			msgbox "认证出错, 请检查密码"
		end if
	'end if
loop