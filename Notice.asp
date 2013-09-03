<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/connect.asp"-->
<%
call aspsql()
id=trim(request("id"))
if id="" then response.redirect "/index.asp"
	Set rsprod=conn.execute ("select * from DU_Notice where  id='"&id&"'")
	if (rsprod.bof and rsprod.eof) then
		response.redirect "../index.asp"
	else
		message_title=rsprod("Ntitle")
		message_body=rsprod("NContent")
		message_NEditTime=rsprod("NEditTime")
%>
<% end if %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="image/web_css/login.css" rel="stylesheet" type="text/css" />
<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
</head>
<body>
<br/>
<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td><strong><center><font color="#FF0000"><%=message_title%></font></center></strong></td>
  </tr>
  <tr>
    <td><%=message_body%></td>
  </tr>
    <tr align="right">
    <td>裁男&nbsp;&nbsp;&nbsp;</td>
  </tr>
    <tr align="right">
    <td><%=format_time(message_NEditTime,4)%>&nbsp;&nbsp;</td>
  </tr>
</table>
<br/>
</body>
</html>
<% call CloseConn() %>
