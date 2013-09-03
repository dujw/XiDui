<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<!--#include file="../../include/md5.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;网站帮助信息设置</font></B></font></TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table width="99%" border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
  <%
Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from dk501_message order by id asc"
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "目前共有 0 条记录"
else
	do while not rs.EOF
	%>
  <tr>
    <td width="35%" height="30" align="right" bgcolor="#f4f5fd">帮助信息设置：<input type=text value="<%=rs("message_title")%>" name=zimu size=30> 
    &nbsp;</td>
    <td width="65%" bgcolor="#f4f5fd">&nbsp;<a href="manage_message_Modify.asp?id=<%=rs("id")%>"><img src="../Images/fav.gif" width="57" height="20" border="0" /></a> 编辑您想要的信息！</td>
  </tr>
<%
rs.movenext
loop
end if
%>
</table>
<br><br>
</BODY>
</html>
<% call CloseConn() %>
