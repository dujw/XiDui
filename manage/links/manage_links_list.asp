<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
</head>
<BODY>
<%
action=request("action")
if action="" then 
%>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER: #8B979F 1px solid; ">
    <TR>
      <TD height="30" bgColor=#f4f5fd><font color="#000000">��<B><font color="#FF0000">&nbsp;������������</font></B></font></TD>
    </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
  <form name=linklist action="manage_links_list.asp" method=post>
    <tr bgcolor="#f4f5fd" > 
      <td width="5%" valign="middle"><strong>&nbsp;ѡ</strong></td>
      <td width="10%" valign="middle"><strong>&nbsp;���</strong></td>
      <td width="40%" height="30" valign="middle" bgcolor="#f4f5fd"><strong>&nbsp;��վ����</strong></td>
      <td width="30%" height="30" valign="middle"><strong>&nbsp;logo</strong></td>
      <td width="15%" height="30" valign="middle"><strong>&nbsp;״̬</strong></td>
    </tr>
    <%
msg_per_page = 20     'ÿҳ��ʾ��¼����

set rs=Server.CreateObject("ADODB.recordset")
sql="select * from dk501_links order by txt desc,num asc,addtime asc"

rs.cursorlocation = 3    '�����α�
rs.pagesize = msg_per_page
rs.open sql,conn,1,1 


if rs.eof and rs.bof then
	response.write "<tr align='center' style='vertical-align: top'><td colspan='5'>��ʱû����������</td></tr>"
else '����¼���Ƿ�Ϊ��
	totalrec = RS.RecordCount 'totalrec���ܼ�¼����
	if rs.recordcount mod msg_per_page = 0 then
	n = rs.recordcount\msg_per_page 'n:��ҳ��
	else 
	n = rs.recordcount\msg_per_page+1 
	end if 

	currentpage = request("page")
	If currentpage <> "" then
		currentpage = clng(currentpage)
		if currentpage < 1 then currentpage = 1
	else
		currentpage = 1
	End if 
	if currentpage*msg_per_page > totalrec and not((currentpage-1)*msg_per_page < totalrec) then 
		currentPage=1
	end if
	rs.absolutepage = currentpage 'absolutepage������ָ��ָ��ĳҳ��ͷ
	rowcount = rs.pagesize   '����ÿһҳ�����ݼ�¼��

	dim i
	dim k

	Do while not rs.eof and rowcount>0
	%>
    <tr bgcolor="#FFFFFF" > 
      <td height="25" class=b> <input type='checkbox' value='<%=Cstr(rs("id"))%>' name=linksid> 
      </td>
      <td height="25"><%=rs("num")%></td>
      <td height="25"><a href="manage_links_list.asp?action=modify_links&linksid=<%=rs("id")%>" alt='�༭������������'><%=rs("site")%></a> 
      </td>
      <td height="25"> <a href="../gotourl.asp?<%=rs("url")%>" target=_blank alt="ȥ�����վ���"> 
        <%
if rs("txt")="1" then
response.write "��������"
else
if left(rs("logo"),7)="http://" then
response.write "<img src='"&rs("logo")&"' border=0 width=88 height=31></a>"
else
response.write "<img src='../"&rs("logo")&"' border=0 width=88 height=31></a>"
end if
end if
%>
      </td>
      <td height="25"> 
        <% if rs("online")="1" then response.write "��ʾ"%>
        <% if rs("online")="0" then response.write "<font color=red>����</font>"%>
      </td>
    </tr>
    <%
	rowcount=rowcount-1
	rs.movenext
	if rs.eof then exit do
	loop	
end if     
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
    <tr bgcolor="#FFFFFF"> 
      <td height="30" colspan=7 bgcolor="#f4f5fd" class=b> <input type='checkbox' name=chkall onclick='CheckAll(this.form)'>
        ȫѡ 
        <input type="submit" name="action" value="�ر�" onClick="{if(confirm('ȷ��Ҫ����ҳ����ѡ����������')){this.document.linklist.submit();return true;}return false;}"> 
        <input type="submit" name="action" value="��" onClick="{if(confirm('ȷ��Ҫ����ҳ��ʾѡ����������')){this.document.linklist.submit();return true;}return false;}"> 
        <input type="submit" name="action" value="ɾ��" onClick="{if(confirm('�ò������ɻָ���\n\nȷ��ɾ��ѡ������Ŀ��')){this.document.linklist.submit();return true;}return false;}"> 
      <input type="button" name="action2" onClick="javascript:location.href='manage_links_list.asp?action=add_links';" value="�����������"></td>
    </tr>
  </form>
</table>
		  <table height="5"><tr><td></td></tr></table>

<%call listPages()%>
<%end if

if action="add_links" then
	call add_links()
end if

if action="ɾ��" then
delid=replace(request("id")," ","")
call delall_links()
end if

if action="�ر�" then
delid=replace(request("id")," ","")
call close_links()
end if

if action="��" then
delid=replace(request("id")," ","")
call open_links()
end if

if action="modify_links" then
	id=replace(request("id")," ","")
	call modify_links()
end if%></table>
</td></tr> </table> 
</body>
</html>

<%
sub listPages() 
if n <= 1 then exit sub 
%>
<p>>> 
<%if currentpage = 1 then%>
<font color=darkgray>��ҳ</font>
<%else%> 
<font color=black face="arial"><a href="<%=request.ServerVariables("script_name")%>?page=1">
��ҳ</font></a>
 <a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage-1%>">
<font color=black face="arial" >ǰҳ</a></font>
<%end if%>
<%if currentpage = n then%> 
<font color=darkgray face="arial" >��ҳ</font>
<%else%> 
<font color=black face="arial" ><a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage+1%>">��ҳ</a>
 <a href="<%=request.ServerVariables("script_name")%>?page=<%=n%>">ĩҳ</a></font>
<%end if%>
<font color=black face="arial" >
  ҳ�Σ�<%=currentpage%>/<%=n%>ҳ&nbsp;&nbsp;ÿҳ����<%=msg_per_page%>&nbsp;&nbsp;������<%=totalrec%></font></p>
<%end sub

sub delall_links()
	if request("linksid")="" or isnull(request("linksid")) then
	response.write "<script language='javascript'>"
	response.write "alert('����ʧ�ܣ���ʲôҲû��ѡ��');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
	else
		conn.execute("delete from dk501_links where Id in ("&request("linksid")&")")
		if err.number<>0 then '������
			response.write "���ݿ����ʧ�ܣ�" & err.description
			err.clear
		else		
		 response.write "<script language=javascript>alert('��ѡ���������ѱ�ɾ����');location.href='manage_links_list.asp';</script>"
		end if
	end if
end sub

sub close_links()       '�ر�
   if request("linksid")="" or isnull(request("linksid")) then
	response.write "<script language='javascript'>"
	response.write "alert('����ʧ�ܣ���ʲôҲû��ѡ��');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
   else
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from dk501_links where Id in ("&request("linksid")&")"
	rs.open sql,conn,1,3
	Do while not rs.eof
	rs("online")="0"
	rs.update
	rs.movenext
	loop			

	response.write "<script language='javascript'>"
	response.write "alert('��ѡ���������ѱ��رա�');"
	response.write "location.href='manage_links_list.asp';"
	response.write "</script>"
   end if
end sub


sub open_links()	'��
   if request("linksid")="" or isnull(request("linksid")) then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ���ʲôҲû��ѡ��');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
   else
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from dk501_links where Id in ("&request("linksid")&")"
	rs.open sql,conn,1,3
	Do while not rs.eof
	rs("online")="1"
	rs.update
	rs.movenext
	loop			

	response.write "<script language='javascript'>"
	response.write "alert('�ɹ�����ѡ���������ѱ��򿪡�');"
	response.write "location.href='manage_links_list.asp';"
	response.write "</script>"
   end if
end sub


sub add_links()
if  request("send")="ok" then

'������д������
if request("site")="" or request("num")="" or request("url")="" then
	response.write "<script language='javascript'>"
	response.write "alert('�������ӵ�������д���������������д����');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end if

if request.form("txt")="0" and request.form("logo")="" then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ���ѡ����ͼ�����ӣ���δ��д��Ч��ͼ���ַ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"	
	response.end
end if

	sql="select * from dk501_links"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3

	rs.addnew
		rs("site")=trim(request.form("site"))
	  	rs("logo")=trim(request.form("logo"))
	  	rs("url")=trim(request.form("url"))
	  	rs("num")=trim(request.form("num"))
	  	rs("online")=request.form("online")
	  	rs("txt")=request.form("txt")

	rs.update	
	rs.close
	set rs=nothing
		 response.write "<script language=javascript>alert('��ϲ���Ѿ������һ���µ��������ӣ�');location.href='manage_links_list.asp';</script>"

else
%>
</p>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;�����������</font></B></font></TD>
  </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
<form name="add" method="post" action="manage_links_list.asp?action=add_links">
    <tr bgcolor="#FFFFFF">
      <td  align=right>��ţ�</td>
      <td >
        <input name="num" type="text"  size="5" maxlength="2"> ������ŵ���Դ�С����</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td  align=right>���ƣ�</td>
      <td >
        <input name="site" type="text"  size="40" maxlength="40"> ��д��վ����</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align=right>ͼ�꣺</td>
      <td>
        <input name="logo" type="text"  size="40" maxlength="100" value="http://"> ��дlogoͼ�����ַ</td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td   align=right>��ַ��</td>
      <td >
        <input name="url" type="text"  size="40" maxlength="100" value="http://"> ��д������վ����ַ</td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td align=right>���ͣ�</td>
      <td>
	<input name="txt" type="radio" value='0' >ͼ������
        <input name="txt" type="radio" value='1' checked>��������  </td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td   align=right>״̬��</td>
      <td >
        <input name="online" type="radio" value="1" checked>��ʾ  
	<input name="online" type="radio" value="0">����</td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td width="10%"  align=right>&nbsp; </td>
      <td width="90%"><input type="submit" name="SubmitName" value=" �� �� ">
          <input name="send" type="hidden"  value="ok"></td>
    </tr>
	</form>
  </table>

<%
end if
end sub

sub modify_links()
if  request.form("send")<>"ok" then
	  set rs=Server.CreateObject("adodb.recordset")
	  sql="select * from dk501_links where id="&request("linksid")
	  rs.open sql,conn,1,3

%>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;�޸�����</font></B></font></TD>
  </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
<form name="modify" method="post" action="manage_links_list.asp?action=modify_links">
    <tr  bgcolor="#FFFFFF">
      <td  align=right>��ţ�</td>
      <td >
        <input name="num" type="text"  size="5" maxlength="2" value=<%=rs("num")%>> ������ŵ���Դ�С����</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td  align=right>���ƣ�</td>
      <td >
        <input name="site" type="text"  size="40" maxlength="40" value=<%=rs("site")%>> ��д��վ����</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td   align=right>ͼ�꣺</td>
      <td >
        <input name="logo" type="text"  size="40" maxlength="100" value=<%=rs("logo")%>> ��дlogoͼ�����ַ</td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td   align=right>��ַ��</td>
      <td >
        <input name="url" type="text"  size="40" maxlength="100" value=<%=rs("url")%>> ��д������վ����ַ</td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td align=right>���ͣ�</td>
      <td>
	<input name="txt" type="radio" value='0'  <% if rs("txt")="0" then%>checked<%end if%>>ͼ������
        <input name="txt" type="radio" value='1' <% if rs("txt")="1" then%>checked<%end if%>>�������� </td> 
    </tr>

    <tr bgcolor="#FFFFFF">
      <td   align=right>״̬��</td>
      <td >
        <input name="online" type="radio" value='1' <% if rs("online")="1" then%>checked<%end if%>>��ʾ  
	<input name="online" type="radio" value='0'  <% if rs("online")="0" then%>checked<%end if%>>����</td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td width="10%"  align=right>&nbsp; </td>
      <td width="90%"><input type="submit" name="SubmitName" value=" ��  �� ">
          <input name="send" type="hidden"  value="ok">
	<input name="linksid" type="hidden" value="<%=request("linksid")%>"></td>
    </tr>
	</form>
  </table>
		
<%
	  rs.close
	  set rs=nothing
else

'������д������
if request.form("site")="" or request.form("num")="" or request.form("url")="" then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ��������ӵ�������д��������');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"	
	response.end
end if

if request.form("txt")="0" and request.form("logo")="" then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ���ѡ����ͼ�����ӣ���δ��д��Ч��ͼ���ַ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"	
	response.end
end if


	  set rs=Server.CreateObject("adodb.recordset")
	  sql="select * from dk501_links where id="&request("linksid")
	  rs.open sql,conn,1,3
	  if rs.eof and rs.bof then
		 response.write "<script language=javascript>alert('��Ŵ���');location.href='manage_links_list.asp';</script>"
		 response.end
	  else
	  	rs("logo")=trim(request.form("logo"))
	  	rs("site")=trim(request.form("site"))
	  	rs("url")=trim(request.form("url"))
	  	rs("num")=trim(request.form("num"))
	  	rs("online")=trim(request.form("online"))
	  	rs("txt")=trim(request.form("txt"))
		rs.update	
	  end if
	  rs.close
	  set rs=nothing
	 call CloseConn()
		 response.write "<script language=javascript>alert('��ϲ�����������Ѿ��޸ĳɹ���');location.href='manage_links_list.asp';</script>"
end if
end sub%>
