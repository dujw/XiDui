<%
'检查用户是否登陆
if request.cookies("buyok")("admin")="" then response.write "<meta http-equiv='refresh' content='0;URL=../login.asp'>"
%>
<%
'检查用户是否是非法登陆和没有登陆操作网站数据
sub checkmanage(str)
	Set mrs = conn.Execute("select manage from manage_user where username='"&request.cookies("buyok")("admin")&"'")
	if not (mrs.bof and mrs.eof) then
	manage=mrs("manage")
		if instr(manage,str)<=0 then
			response.write "<script language='javascript'>"
			response.write "alert('警告：您没有此项操作的权限！');"
			response.write "location.href='/manage/quit.asp';"
			response.write "</script>"
			response.end
		else
			session("buyok_admin_login")=0
		end if
	else 
		response.write "<script language='javascript'>"
		response.write "alert('没有登陆，不能执行此操作！');"
		response.write "location.href='/manage/quit.asp';"
		response.write "</script>"
		response.end
	end if
	set mrs=nothing
end sub
%>

