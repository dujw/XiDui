<!--#include file="../include/connect.asp"-->
<!--#include file="../include/md5.asp"-->
<%
call aspsql()
Userid=trim(request.form("userid"))
Password=trim(request.form("password"))
'check1=trim(request("check1"))   '输入的验证码
'check2=trim(request("check2")) '系统自带的验证码
'if check1="" or check1 <> check2 then	
'	Response.Write("<script language=""JavaScript"">alert(""错误：请输入正确的验证码，请返回检查！！"");history.go(-1);/script>")
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
	response.write "alert('出错了：没有这个用户 请重新登陆。');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	'response.redirect "error.asp?error=003"
	response.end
end if
if rs("Status")<>"正常" and rs("Status")<>"1" then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您的会员号已被锁定或者未通过审核。');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
if  rs("UserPassword")<> md5(Password) then
	session("login_error")=session("login_error")+1
	response.write "<script language='javascript'>"
	response.write "alert('您输入的密码不正确，请检查后重新输入。\n\n出错 "&session("login_error")&" 次');"
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