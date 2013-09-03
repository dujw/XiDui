<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<%
if Request.QueryString("mark")="southidc" then
		ntitle=Request("title")
		novertime=trim(Request("overtime"))
		ncontent=trim(Request("content"))
		nistop=trim(Request("istop"))
		If novertime="" Then
			Response.write  "<script language=javascript>alert('错误!请输入过期时间!!');history.go(-1);</script>"
			Response.end
		end if	
			
		If ncontent="" Then
			Response.write  "<script language=javascript>alert('错误!请输入内容!!');history.go(-1);</script>"
			Response.end
		end if			
			
		
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from DU_Notice"
	rs.open sql,conn,1,3
	rs.addnew
	rs("Ntitle")=ntitle	
	rs("NOverTime")=novertime
	rs("NContent")=ncontent
	IF nistop<>"" THEN 
		rs("NIsTop")=nistop
	END IF
	rs("NAddUser")=request.cookies("buyok")("admin")
	rs.update
	rs.close
	Response.write  "<script language=javascript>alert('添加成功!!');window.location.href='manage_notice.asp'</script>"
	'response.redirect "manage_notice.asp"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../../js_lib/Calendar.js"></script>
</head>
<BODY>
<form method=POST action="?mark=southidc">
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;网站通知发布</font></B></font></TD>
  </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="99%" border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
  <tr>
    <td width="11%" height="25" align="right" bgcolor="#f4f5fd">通知标题：</td>
    <td width="89%" height="25" bgcolor="#f4f5fd"><input name="title" type="text" id="title" value=""  width="260px;">&nbsp;&nbsp;&nbsp;&nbsp; <input name="istop" type="checkbox" value="1">置顶</td>
  </tr>
    <tr>
    <td width="11%" height="25" align="right" bgcolor="#f4f5fd">过期时间：</td>
    <td width="89%" height="25" bgcolor="#f4f5fd"><input name="overtime" type="text" id="overtime" width="260px;" onFocus="calendar()">
    (正确选择过期时间，超过过期时间的通知不会显示)</td>
  </tr>
  <tr>
    <td height="25" align="right" bgcolor="#f4f5fd">通知内容：</td>
    <td width="89%" height="25" bgcolor="#f4f5fd">
	<input type="hidden" name="content" value=""> 
     <iframe ID="eWebEditor1" src="../Southidceditor/ewebeditor.asp?id=content&style=southidc" frameborder="0" scrolling="no" width="700" HEIGHT="300"></iframe>	
	</td>
  </tr>
  <tr>
    <td height="25" align="right" bgcolor="#f4f5fd"></td>
    <td width="89%" height="25" bgcolor="#f4f5fd"><input type="submit" name="Submit" value="提交添加" /><input type="reset" name="Submit" value="内容全部重置" /></td>
  </tr>
	  
</table></form>
<br>
</BODY>
</html>
<% call CloseConn() %>
