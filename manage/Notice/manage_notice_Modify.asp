<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<% id=trim(request("ID"))
	if Request.QueryString("mark")="southidc" then
		id=trim(request("id"))
		ntitle=Request("title")
		novertime=trim(Request("overtime"))
		ncontent=trim(Request("content"))
		nistop=trim(Request("istop"))
		If novertime="" Then
			Response.write  "<script language=javascript>alert('����!���������ʱ��!!');history.go(-1);</script>"
			Response.end
		end if	
			
		If ncontent="" Then
			Response.write  "<script language=javascript>alert('����!����������!!');history.go(-1);</script>"
			Response.end
		end if			
			
		
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from DU_Notice where id="&id 
	rs.open sql,conn,1,3
	
		if rs.eof then
				Response.write  "<script language=javascript>alert('����!·����������ϵ����Ա!!');history.go(-1);</script>"
				Response.end
		else			
	rs("Ntitle")=ntitle	
	rs("NOverTime")=novertime
	rs("NContent")=ncontent
	IF nistop<>"" THEN 
		rs("NIsTop")=nistop
		ELSE
		rs("NIsTop")=0
	END IF
	rs("NAddUser")=request.cookies("buyok")("admin")
	rs("NEditTime")=NOW()
	rs.update
	rs.close
	Response.write  "<script language=javascript>alert('�޸ĳɹ�!!');window.location.href='manage_notice.asp'</script>"
	end if
	END IF
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
<%
id=trim(request("ID"))
Set rabout=Server.CreateObject("ADODB.RecordSet") 
sql="select * from DU_Notice where id="&id
rabout.Open sql,conn,1,1
	if rabout.eof and rabout.bof then
		Response.write  "<script language=javascript>alert('����!�Բ������ݿ����!!');history.go(-1);</script>"
		Response.End
	else
%>

 <input type=hidden name=id value=<%=id%>>
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;��վ֪ͨ�޸�</font></B></font></TD>
  </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="99%" border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
  <tr>
    <td width="11%" height="25" align="right" bgcolor="#f4f5fd">֪ͨ���⣺</td>
    <td width="89%" height="25" bgcolor="#f4f5fd"><input name="title" type="text" id="title" value="<%=rabout("Ntitle")%>"  width="260px;">&nbsp;&nbsp;&nbsp;&nbsp; <input name="istop" type="checkbox" value="1" <%if rabout("NIsTop") then %>checked<%end if%>>�ö�</td>
  </tr>
    <tr>
    <td width="11%" height="25" align="right" bgcolor="#f4f5fd">����ʱ�䣺</td>
    <td width="89%" height="25" bgcolor="#f4f5fd"><input name="overtime" type="text" id="overtime" value="<%=rabout("NOverTime")%>"  width="260px;"  onFocus="calendar()">
      (��ȷѡ�����ʱ�䣬��������ʱ���֪ͨ������ʾ)</td>
  </tr>
  <tr>
    <td height="25" align="right" bgcolor="#f4f5fd">֪ͨ���ݣ�</td>
    <td width="89%" height="25" bgcolor="#f4f5fd">
	<input type="hidden" name="content" value="<%=Server.HTMLEncode(rabout("NContent"))%>"> 
     <iframe ID="eWebEditor1" src="../Southidceditor/ewebeditor.asp?id=content&style=southidc" frameborder="0" scrolling="no" width="700" HEIGHT="300"></iframe>	
	</td>
  </tr>
  <tr>
    <td height="25" align="right" bgcolor="#f4f5fd"></td>
    <td width="89%" height="25" bgcolor="#f4f5fd"><input type="submit" name="Submit" value="�ύ�޸�" /><input type="reset" name="Submit" value="����ȫ������" /></td>
  </tr>
  <% 
		end if		
		rabout.close
		set rabout=nothing
	  %>
	  
</table></form>
<br>
</BODY>
</html>
<% call CloseConn() %>
