<!--#include file="../include/connect.asp"-->
<!--#include file="../include/md5.asp"-->
<%
call aspsql()
Userid=trim(request.form("userid"))
Password=trim(request.form("password"))
'check1=trim(request("check1"))   '�������֤��
'check2=trim(request("check2")) 'ϵͳ�Դ�����֤��
'if check1="" or check1 <> check2 then	
'	Response.Write("<script language=""JavaScript"">alert(""������������ȷ����֤�룬�뷵�ؼ�飡��"");history.go(-1);/script>")
'end if

if Userid="" or Password =""  then	
	response.redirect "/error/error.asp?error=004"
	response.End()
end if 
if Userid = request.cookies("buyok")("userid") then 
	response.redirect "/error/error.asp?error=005"
	response.End()
end if 


if request.form("Login")<>"ok" then response.redirect "/index.asp"
sql = "select lastlogin,IP,TotalLogin,Status,UserPassword,userid from dk501_user where userid='"&Userid&"'"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,3
if (rs.bof and rs.eof) then	
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�û������û� �����µ�½��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	'response.redirect "error.asp?error=003"
	response.end
end if
if rs("Status")<>"����" and rs("Status")<>"1" then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ����Ļ�Ա���ѱ���������δͨ����ˡ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
if  rs("UserPassword")<> md5(Password) then
	session("login_error")=session("login_error")+1
	response.write "<script language='javascript'>"
	response.write "alert('����������벻��ȷ��������������롣\n\n���� "&session("login_error")&" ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
else
	rs("lastlogin")=now()
	rs("IP")=Request.serverVariables("REMOTE_ADDR")
	rs("TotalLogin")=rs("TotalLogin")+1
	rs.update
	response.cookies("buyok")("userid")=lcase(userid)
	if request.form("cook")<>"0" then response.cookies("buyok").expires=now+cook
	response.write "<meta http-equiv='refresh' content='0;URL=/user/index.asp'>"
end if
rs.close
set rs=nothing
call CloseConn()
%>