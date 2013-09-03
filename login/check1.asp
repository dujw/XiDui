<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../include/connect.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="../image/web_css/login.css" rel="stylesheet" type="text/css" />
<script language="JavaScript">
function closeit() {
setTimeout("self.close()",8000)
}
</script>
</HEAD>
<BODY leftmargin="0" topmargin="0" background=images/admin/back.gif onLoad="closeit()">
<%
call aspsql()
userid=Trim(request.ServerVariables("query_string"))
if userid = "" then
	    response.write "<img border=0 src=../image/login/wrong.gif alt='出错了，您没有输入用户名，或者输入的用户名中含有非法字符'>"
	    response.end
elseif buyoktxtcheck(userid)<>userid then
	    response.write "<img border=0 src=images/small/wrong.gif alt='出错了，您没有输入用户名，或者输入的用户名中含有非法字符'>"
	    response.end
else
    set rs = conn.execute("SELECT UserId FROM dk501_user where UserId= '" & UserId & "'")
    if not (rs.Bof or rs.eof) then
	response.write "<img border=0 src=../image/login/no.gif alt='非常遗憾，此用户名已被他人注册，请选用其它用户名。'>"
    else
 	response.write "<img border=0 src=../image/login/yes.gif alt='恭喜，您可以使用此用户名。'>"
    end if
    set rs=nothing
end if

conn.close
set conn=nothing
%>
</BODY>
</HTML>
