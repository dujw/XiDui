<%
'����û��Ƿ��½
if request.cookies("buyok")("admin")="" then response.write "<meta http-equiv='refresh' content='0;URL=../login.asp'>"
%>
<%
'����û��Ƿ��ǷǷ���½��û�е�½������վ����
sub checkmanage(str)
	Set mrs = conn.Execute("select manage from manage_user where username='"&request.cookies("buyok")("admin")&"'")
	if not (mrs.bof and mrs.eof) then
	manage=mrs("manage")
		if instr(manage,str)<=0 then
			response.write "<script language='javascript'>"
			response.write "alert('���棺��û�д��������Ȩ�ޣ�');"
			response.write "location.href='/manage/quit.asp';"
			response.write "</script>"
			response.end
		else
			session("buyok_admin_login")=0
		end if
	else 
		response.write "<script language='javascript'>"
		response.write "alert('û�е�½������ִ�д˲�����');"
		response.write "location.href='/manage/quit.asp';"
		response.write "</script>"
		response.end
	end if
	set mrs=nothing
end sub
%>

