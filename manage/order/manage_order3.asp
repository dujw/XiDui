<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
<script language=javascript>
<!--
function CheckAll(form){
for (var i=0;i<form.elements.length;i++){
var e = form.elements[i];
if (e.name != 'chkall') e.checked = form.chkall.checked; 
}
}
-->
</script>
</head>
<BODY>

<%
action = request("action")
if action="" then call order_index()
if action="�ָ�����" then call huifu()
if action="����ɾ��" then call orderdel()
sub order_index()
%>

<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">��<B><font color="#FF0000">&nbsp;��������վ</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>

<table width="99%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
<form action=manage_order3.asp method=post name=orderlist>
<%pages = 1000
set rs = server.createobject("adodb.recordset")
sql = "select a.ID,a.OrderNum,a.del,a.OrderSum,a.UserId,a.RecName2,a.OrderTime,a.Status,b.StatusDefines from dk501_borderList a,dk501_types b where a.del=1 and a.Status=b.Status "
sql= sql+" order by OrderTime desc"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
response.write "<tr bgcolor=#FFFFFF><td colspan=7 align=center height=50>����վ�ǿյ�</td></tr>"
response.end 
rs.close
set rs=nothing
end if
totalrs=rs.RecordCount
rs.pageSize = pages
allPages = rs.pageCount
page = Request("page")
If not isNumeric(page) then page=1
if isEmpty(page) or int(page) <= 1 then
page = 1
elseif int(page) > allPages then
page = allPages 
end if
rs.AbsolutePage = page
%>

<tr bgcolor="#f4f5fd">
<td width=5% height="25" align=center>ѡ</td>
<td width='15%' height="25" align=center>������</td>
<td width='15%' height="25" align=center>���</td>
<td width='15%' height="25" align=center>��ԱID</td>
<td width='12%' height="25" align=center>�ջ�������</td>
<td width='20%' height="25" align=center>�µ�ʱ��</td>
<td width='18%' height="25" align=center>����״̬</td>
</tr>
<%
Do While Not rs.eof and pages>0
%>
<tr bgcolor=#FFFFFF><td><input type='checkbox' value=<%=rs("ID")%> name=id></td>
<td align=center><%=rs("OrderNum")%></td>
<td align=center><%=rs("OrderSum")%></td>
<td><%=rs("UserId")%></td>
<td><%=rs("RecName2")%></td>
<td><%=rs("OrderTime")%></td>
<td align=center>
<%
if rs("Status")="0" then
	response.write "<font color=red>"&rs("StatusDefines")&"</font>"
elseif rs("Status")="99" then 
	response.write "<font color=blue>"&rs("StatusDefines")&"</font>"
elseif rs("Status")="88" then 
	response.write "<font color=blue>"&rs("StatusDefines")&"</font>"
else
	response.write "<font color=darkgray>"&rs("StatusDefines")&"</font>"
end if
response.write "</td></tr>"
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
%>	
<tr bgcolor="#FFFFFF"><td colspan=7>
<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>
<input type="submit" name="action" value="����ɾ��" onClick="{if(confirm('��ʾ���ò������ɻָ�����ȷʵҪɾ����ѡ������')){this.document.orderlist.submit();return true;}return false;}"> 	
<input type="submit" name="action" value="�ָ�����" onClick="{if(confirm('��ʾ��ȷ��Ҫ�ָ��˶�����')){this.document.orderlist.submit();return true;}return false;}"> 	
</td></tr>
</form>
</table>
		  <table height="5"><tr><td></td></tr></table>
<%
rs.close
set rs=nothing
'if allpages<=1 then exit sub
response.write "<br>�ܶ�����<font color=red>"&totalrs&"</font> &nbsp; ÿҳ<font color=red>20</font>&nbsp; "
if page = 1 then
response.write "<font color=darkgray>��ҳ ǰҳ</font>"
else
response.write "<a href=manage_order3.asp?keywords="&keywords&"&page=1>��ҳ</a> <a href=manage_order3.asp?keywords="&keywords&"&page="&page-1&">ǰҳ</a>"
end if
if page = allpages then
response.write "<font color=darkgray> ��ҳ ĩҳ</font>"
else
response.write " <a href=manage_order3.asp?keywords="&keywords&"&page="&page+1&">��ҳ</a> <a href=manage_order3.asp?keywords="&keywords&"&page="&allpages&">ĩҳ</a>"
end if

end sub



sub orderdel()
	delid=request("id")
	if Delid="" or isnull(Delid) then
	response.write "<script language='javascript'>"
	response.write "alert('����ʧ�ܣ���ʲôҲû��ѡ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if
	
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select id,OrderNum from dk501_borderList where id in ("&Delid&")"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		response.write "<script language='javascript'>"
		response.write "alert('�����ˣ��ö��������ڻ����Ѿ���ɾ����');"		
		response.write "</script>"			
		response.end 
	else
		Do while not rs.eof
			OrderNum = rs("OrderNum")
			Set rs_border=Server.CreateObject("ADODB.Recordset")
			sql_border="delete  from dk501_border where OrderNum = '"&rs("OrderNum")&"'"
			rs_border.open sql_border,conn,1,3
		rs.delete
		rs.update
		rs.movenext
		loop
		rs.close
		set rs=nothing
		response.write "<script language='javascript'>"
		response.write "alert('�����ɹ�����ѡ�����ѱ�����ɾ����');"
		response.write "location.href='manage_order3.asp';"			
		response.write "</script>"
	end if
end sub



sub huifu()
id=request("id")
if id="" or isnull(id) then
response.write "<script language='javascript'>"
response.write "alert('����ʧ�ܣ���ʲôҲû��ѡ��');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from dk501_borderList where id in ("&id&")"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ��ö��������ڻ����Ѿ���ɾ����');"		
response.write "</script>"			
response.end 
else
Do while not rs.eof
rs("del")=false
rs.update
rs.movenext
loop
rs.close
set rs=nothing
response.write "<script language='javascript'>"
response.write "alert('�����ɹ�����ѡ�����ѱ��ָ���');"
response.write "location.href='manage_order3.asp';"			
response.write "</script>"
end if
end sub

%>
</body>
</html>