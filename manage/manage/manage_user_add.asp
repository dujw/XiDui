<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<!--#include file="../../include/md5.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<%
action=request("add")
if action<>"ok" then
%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid">
    <TR> 
      <TD height="30" bgColor=#f4f5fd><font color="#000000">　<B><font color="#FF0000">添加管理员</font></B></font></TD>
    </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>

<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
  <form action=manage_user_add.asp method=post>
    <tr bgcolor="#FFFFFF">
      <td width=20% align=right height="30">管理员名称：</td>
      <td height="30"><input type="text" name="Username" size="20" value="" maxlength=16> 
        &nbsp;&nbsp; <img src=../images/admin/memo.gif alt="用户名长度：6-16位，可使用中文和字母"> 
      </td>
    </TR>
    <tr bgcolor="#FFFFFF">
      <td width=20% align=right height="30">请输入密码：</td>
      <td height="30"><input type="password" name="Pass1" size="20"  maxlength=16> &nbsp;&nbsp; 
      <img src=../images/admin/memo.gif alt="密码长度：8-16位<br>建议使用数字和字母组合"> </td>
    </TR>
    <tr bgcolor="#FFFFFF">
      <td width=20% align=right height="30">请确认密码：</td>
      <td height="30"><input type="password" name="Pass2" size="20"  maxlength=16> &nbsp;&nbsp; 
      <img src=../images/admin/memo.gif alt="请确认密码"> </td>
    </TR>
    <tr bgcolor="#FFFFFF">
      <td height="30" colspan=2><input type="submit" name="Submit" value="确认添加"> <input type="hidden" name="add" value="ok"></td>
    </tr>
  </form>
</table>

<%   
else
	if trim(Request("Username"))="" or trim(Request("pass1"))="" or trim(Request("pass2"))="" then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，填写不完整，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if llen(trim(Request("Username")))<4 or llen(trim(Request("Username")))>16 then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您输入的管理员名称不合要求，要求长度4-16位');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if llen(trim(Request("pass1")))<4 or llen(trim(Request("pass1")))>16 then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您输入的密码长度不合要求，要求长度4-16位');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if trim(Request("pass1"))<>trim(Request("pass2")) then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，两次输入的密码不符，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	Set rs = conn.Execute("select * from manage_user where username='"&request.form("username")&"'")
	if not (rs.eof and rs.bof) then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，该管理员已经存在，请选择其它名称！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from manage_user"
	rs.open sql,conn,1,3
		rs.addnew
		rs("username")=Request.form("Username")
		rs("password")=md5(Request("pass1"))
		rs("manage")="|01|02|03|04|05|06|07|08|09|10|11|12"
		rs.update
		rs.close
		set rs=nothing

	response.write "<script language='javascript'>"
	response.write "alert('操作成功，您已经成功添加一个管理员。\n\n接下来，请为该管理员设置权限！');"
	response.write "location.href='edit_manage_user.asp?admin="&Request.form("Username")&"';"
	response.write "</script>"

end if

%>
<% call CloseConn() %>