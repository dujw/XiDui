<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<% id=trim(request("ID"))
	if Request.QueryString("mark")="southidc" then
		id=trim(request("id"))
		content=Request("content")	
			
		If content="" Then
			Response.write  "<script language=javascript>alert('错误!请输入内容!!');history.go(-1);</script>"
			Response.end
		end if			
			
		
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from dk501_message where id="&id 
	rs.open sql,conn,1,3
	
		if rs.eof then
				Response.write  "<script language=javascript>alert('错误!路径出错。请联系管理员!!');history.go(-1);</script>"
				Response.end
		else			
			rs("message_body")=content
			rs.update
			rs.close
			response.redirect "manage_message.asp"
		end if
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>

<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
<%
id=trim(request("ID"))
Set rabout=Server.CreateObject("ADODB.RecordSet") 
sql="select * from dk501_message where id="&id
rabout.Open sql,conn,1,1
	if rabout.eof and rabout.bof then
		Response.write  "<script language=javascript>alert('错误!对不起，数据库出错。!!');history.go(-1);</script>"
		Response.End
	else
%>
 <form method=POST action="?mark=southidc">
 <input type=hidden name=id value=<%=id%>>
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;网站帮助信息修改设置</font></B></font></TD>
  </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="99%" border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
  <tr>
    <td width="11%" height="25" align="right" bgcolor="#f4f5fd">栏目名称：</td>
    <td width="89%" height="25" bgcolor="#f4f5fd"><input name="title" type="text" id="title" value="<%=rabout("message_title")%>"></td>
  </tr>
  <tr>
    <td height="25" align="right" bgcolor="#f4f5fd">详细内容：</td>
    <td width="89%" height="25" bgcolor="#f4f5fd">
	<input type="hidden" name="content" value="<%=Server.HTMLEncode(rabout("message_body"))%>"> 
     <iframe ID="eWebEditor1" src="../Southidceditor/ewebeditor.asp?id=content&style=southidc" frameborder="0" scrolling="no" width="700" HEIGHT="500"></iframe>	
	</td>
  </tr>
  <tr>
    <td height="25" align="right" bgcolor="#f4f5fd"></td>
    <td width="89%" height="25" bgcolor="#f4f5fd"><input type="submit" name="Submit" value="修改" /><input type="reset" name="Submit" value="重制" /></td>
  </tr>
  <% 
		end if		
		rabout.close
		set rabout=nothing
	  %>
	  </form>
</table>
<br>
</BODY>
</html>
<% call CloseConn() %>
