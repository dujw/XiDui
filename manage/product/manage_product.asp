<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<%
action = request("action")
if action="" then
	call showclass()
end if
'�༭�������
if action="editclass" then
	call editclass()
end if
'ɾ����������
if action="delMidclass" then
	MidCode=request("MidCode")
	LarCode=request("LarCode")
	call delmidclass()
end if
'���Ӷ�������
if action="addMidclass" then
call addmidclass()
end if

if action="addLarclass" then
'����һ������
call addlarclass()
end if
'ɾ��һ������
if action="delLarclass" then
	LarCode=request("LarCode")
	call dellarclass()
end if


sub addmidclass()
'���Ӷ�������
LarCode=trim(request("LarCode"))
if request("add")="ok" then
newclass=trim(request("newclass"))
If trim(newclass)="" or instr(newclass,"&")>0 or instr(newclass,"%")>0 or instr(newclass,"'")>0 or instr(newclass,"&quot;")>0 then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ�������д�������򲻷���Ҫ������������ύ��');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
response.end
end if
set rs=conn.execute("select * from dk501_productclass where MidCode='"&newclass&"' and LarCode='"&LarCode&"'")
if not (rs.eof and rs.bof) then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ��Ѿ���һ��ͬ������������ڣ���ʹ���������ưɡ�');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
set rs=nothing
set rs=nothing
sql="select * from dk501_productclass where LarCode='"&LarCode&"' order by MidSeq"
set rs=conn.execute(sql)
while Not rs.eof
counter=rs("MidSeq")+1
larseq=rs("LarSeq")
rs.movenext
wend
set rsadd=server.createobject("adodb.recordset")
rsadd.open "dk501_productclass",conn,1,3
rsadd.addnew
rsadd("LarSeq")=larseq
rsadd("LarCode")=LarCode
rsadd("MidSeq")=counter
rsadd("MidCode")=newclass
rsadd.update
rsadd.close
set rsadd=nothing
rs.close
set rs=nothing
response.write "<script language='javascript'>"
response.write "alert('�����ɹ���������˶������ࡰ"&trim(request("newclass"))&"����');"
response.write "location.href='manage_product.asp';"
response.write "</script>"
response.end
else
%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">��<B><font color="#FF0000">&nbsp;��Ʒ�������</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
		  <table height="5"><tr><td></td></tr></table>

<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
<tr bgcolor="#FFFFFF"><td height=25 colspan=2 class=td>&nbsp;��Ӷ�������</td>
</tr>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center>һ����������</td><td><%=LarCode%></td></tr>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center valign=top">���ж�����������</td><td>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from dk501_productclass where LarCode='"&LarCode&"' order by MidSeq asc"
rs.open sql,conn,1,1
if not (rs.eof and rs.bof) then
do while not rs.eof
response.write rs("midcode")&"<br>"
rs.movenext
loop
end if
%>
</td></tr>
<form action='manage_product.asp?action=addMidclass' method=post name=addmidclass>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center>����������������</td><td><INPUT TYPE='text' NAME='newclass' size=20 maxlength=20 value=''>[��������<font color=red>% &quot; ' &amp; + = \</font>�Ƚ����ַ�]</td></tr>
<tr bgcolor="#FFFFFF"><td colspan=2><INPUT TYPE='hidden' name=add value='ok'><INPUT TYPE='hidden' name=LarCode value='<%=LarCode%>'><INPUT TYPE='submit' value='��������'>
</td></tr>
</form>
</table>
<table height="5"><tr><td></td></tr></table>

<%
rs.close
set rs=nothing
end if
end sub

sub addlarclass()
'����һ������
if request("add")="ok" then
	If trim(request("newclass"))="" or instr(request("newclass"),"&")>0 or instr(request("newclass"),"%")>0 or instr(request("newclass"),"'")>0 or instr(request("newclass"),"&quot;")>0 then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�������д�������򲻷���Ҫ������������ύ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if
	set rs=conn.execute("select * from dk501_productclass where LarCode='"&trim(request("newclass"))&"'")
		if not (rs.eof and rs.bof) then
			response.write "<script language='javascript'>"
			response.write "alert('�����ˣ��Ѿ���һ��ͬ��������ڣ���ʹ���������ơ�');"
			response.write "location.href='javascript:history.go(-1)';"
			response.write "</script>"
			response.end
		end if
	set rs=nothing
	set rs=conn.execute("select * from dk501_productclass order by larseq desc")
	if not (rs.eof and rs.bof) then
	count=clng(rs("larseq"))+1
else
	count=1
end if
set rs=nothing
set rsadd=Server.CreateObject("ADODB.Recordset")
sql="SELECT * FROM dk501_productclass"
rsadd.open sql,conn,3,3
rsadd.addnew
rsadd("LarSeq")=count
rsadd("MidSeq")=1
rsadd("LarCode")=trim(request.form("newclass"))
rsadd("MidCode")=trim(request.form("newclass"))
rsadd.update
rsadd.close
set rsadd=nothing
response.write "<script language='javascript'>"
response.write "alert('�����ɹ��������һ�����ࡰ"&trim(request.form("newclass"))&"������һ��ͬ���Ķ������ࡣ');"
response.write "location.href='manage_product.asp';"
response.write "</script>"
response.end
else
set rs=server.createobject("adodb.recordset")
sql="select distinct LarCode,LarSeq from dk501_productclass order by LarSeq asc"
rs.open sql,conn,1,1
%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid">
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">��<B><font color="#FF0000">&nbsp;��Ʒ�������</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>

<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
<form action='manage_product.asp?action=addLarclass' method=post name=addlarclass>
<tr bgcolor="#FFFFFF"><td height=25 colspan=2 class=td>&nbsp;���һ������</td>
</tr>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center>һ����������</td><td><INPUT TYPE='text' NAME='newclass' size=20 maxlength=20 value=''>[��������<font color=red>% &quot; ' &amp; + = \</font>�Ƚ����ַ�]</td></tr>
<tr bgcolor="#FFFFFF"><td colspan=2><INPUT TYPE='hidden' name=add value='ok'><INPUT TYPE='submit' value='��������'>
</td></tr>
</form>
</table>
<%
rs.close
set rs=nothing
end if
end sub


sub delmidclass()
'ɾ����������
conn.execute "delete from dk501_productclass where MidCode='"&request("MidCode")&"' and  larCode='"&request("larCode")&"'" 
'ͬ��ɾ���÷����µ�������Ʒ
if tree_del_yesorno="1" then
conn.execute "delete from dk501_product where MidCode='"&request("MidCode")&"' and  larCode='"&request("larCode")&"'" 
else
conn.execute "update dk501_product set MidCode='δ����',Larcode='δ����' where MidCode='"&request("MidCode")&"' and  larCode='"&request("larCode")&"'" 
end if

conn.close
response.write "<script language='javascript'>"
if tree_del_yesorno="1" then
response.write "alert('�����ɹ�����ѡ���������ѱ�ɾ��\n\nע�⣺�÷����µ�������Ʒ�ѱ�ͬʱɾ����');"
else
response.write "alert('�����ɹ�����ѡ���������ѱ�ɾ��\n\nע�⣺�÷����µ�������Ʒ�ѱ���Ϊδ������Ʒ��');"
end if
response.write "location.href='manage_product.asp';"	
response.write "</script>"
end sub

sub delLarclass()
'ɾ��һ������
conn.execute "delete from dk501_productclass where LarCode='"&request("LarCode")&"'" 
'ͬ��ɾ���÷����µ�������Ʒ
if tree_del_yesorno="1" then
conn.execute "delete from dk501_product where LarCode='"&request("LarCode")&"'" 
else
conn.execute "update dk501_product set MidCode='δ����',Larcode='δ����' where larCode='"&request("larCode")&"'" 
end if


conn.close
response.write "<script language='javascript'>"
if tree_del_yesorno="1" then
response.write "alert('�����ɹ�����ѡһ�������ѱ�ɾ��\n\nע�⣺�÷����µ�������Ʒ�ѱ�ͬʱɾ����');"
else
response.write "alert('�����ɹ�����ѡһ�������ѱ�ɾ��\n\nע�⣺�÷����µ�������Ʒ�ѱ���Ϊδ������Ʒ��');"
end if
response.write "location.href='manage_product.asp';"	
response.write "</script>"
response.end
end sub

sub editclass()
'�༭��Ʒ����
if request("MidCode")<>"" then
 sql="select * from dk501_productclass where MidCode='"&request("MidCode")&"' and  larCode='"&request("larCode")&"'"
 text="MidCode"
elseif request("LarCode")<>"" then
 sql="select * from dk501_productclass where LarCode='"&request("LarCode")&"'"
 text="LarCode"
end if

if request("edit")="ok" then
'�޸Ķ�����������
if request("MidCode")<>"" then

'�������������ͬ��һ���������棬�Ƿ��Ѿ���������ƵĶ����������
set rs=conn.execute("select * from dk501_productclass where MidCode='"&request("MidCode")&"' and  larCode='"&request("larCode")&"'")
if request("old")=request("midcode") then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ��¾�������ͬ��δ���޸ġ�');"
	response.write "location.href='manage_product.asp';"
	response.write "</script>"
	response.end
end if

if not ( rs.eof and rs.bof ) then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ��Ѿ���������ƵĶ������࣬��ʹ���������ơ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
set rs=nothing
'�����������
conn.execute("update dk501_productclass set MidCode='"&trim(request("MidCode"))&"' where Midcode='"&request("old")&"' and LarCode='"&request("larcode")&"'")
'ͬ��������Ʒ�����е����
conn.execute("update dk501_product set MidCode='"&trim(request("MidCode"))&"' where Midcode='"&request("old")&"' and LarCode='"&request("larcode")&"'")


'�޸�һ����������
elseif request("LarCode")<>"" then

set rs=conn.execute("select * from dk501_productclass where larCode='"&request("larCode")&"'")
if request("old")=request("larcode") then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ��¾�������ͬ��δ���޸ġ�');"
	response.write "location.href='manage_product.asp';"
	response.write "</script>"
	response.end
end if

'������������Ƿ��Ѿ���ͬ����һ���������
if not ( rs.eof and rs.bof ) then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ��Ѿ���������Ƶ�һ�����࣬��ʹ���������ơ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
set rs=nothing

'�����������
conn.execute("update dk501_productclass set LarCode='"&trim(request("LarCode"))&"' where LarCode='"&request("old")&"'")
'ͬ��������Ʒ�����е����
conn.execute("update dk501_product set LarCode='"&trim(request("LarCode"))&"' where LarCode='"&request("old")&"'")

end if
conn.close

response.write "<script language='javascript'>"
response.write "alert('�����ɹ�����Ʒ����ɡ�"&request("old")&"������Ϊ��"&request(text)&"��');"
response.write "location.href='manage_product.asp';"
response.write "</script>"
response.end
else
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if text="MidCode" then
%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">��<B><font color="#FF0000">&nbsp;��Ʒ�������</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
		  <table height="5"><tr><td></td></tr></table>

<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
<form name=editform action='manage_product.asp?action=editclass' method=post>
<tr bgcolor="#FFFFFF"><td height=25 colspan=2 class=td>&nbsp;�༭������������</td>
</tr>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center>һ����������</td><td> <%=request("larCode")%> </td></tr>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center>������������</td><td><INPUT TYPE='text' NAME=MidCode size=20 maxlength=20 value='<%=rs("MidCode")%>'>[��������<font color=red>% &quot; ' &amp; + = \</font>�Ƚ����ַ�]</td></tr>
<tr bgcolor="#FFFFFF"><td colspan=2><INPUT TYPE='hidden' name=old value='<%=rs("MidCode")%>'><INPUT TYPE='hidden' name="larcode" value='<%=rs("larcode")%>'><INPUT TYPE='hidden' name=edit value=ok><INPUT TYPE='submit' name=submit value='�����޸�'>
</td></tr>
</form>
</table>
<%
end if
if text="LarCode" then
%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">��<B><font color="#FF0000">&nbsp;��Ʒ��������</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
		  <table height="5"><tr><td></td></tr></table>

<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
<form name=editform action='manage_product.asp?action=editclass' method=post>
<tr bgcolor="#FFFFFF"><td height=25 colspan=2 class=td>&nbsp;�༭һ����������</td>
</tr>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center>һ����������</td><td><INPUT TYPE='text' NAME=LarCode size=20 maxlength=20 value='<%=rs("LarCode")%>'>[��������<font color=red>% &quot; ' &amp; + = \</font>�Ƚ����ַ�]</td></tr>
<tr bgcolor="#FFFFFF"><td colspan=2><INPUT TYPE='hidden' name=old value='<%=rs("LarCode")%>'><INPUT TYPE='hidden' name=edit value=ok><INPUT TYPE='submit' name=submit value='�ύ'>
</td></tr>
</form>
</table>
<%
end if
rs.close
set rs=nothing
end if
end sub


sub showclass()
'���������ҳ
%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">��<B><font color="#FF0000">&nbsp;��Ʒ�������</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30><font color="#000000">&nbsp;</font></TD>
    </TR>
  </TBODY>
</TABLE>
		  <table height="5"><tr><td></td></tr></table>

<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
<%
set rs=server.createobject("adodb.recordset")
sql="select * from dk501_productclass order by LarSeq asc, MidSeq asc"
rs.open sql,conn,1,3
dim old
if rs.eof and rs.bof then 
response.Redirect "manage_product.asp?action=addLarclass"
else
Do While Not rs.eof
if rs("LarCode")<>old then    '�ж���һ��һ�����࣬old��ֵ�ڶ��������б��ֵ
%>
<tr><td width=50% height=30 bgcolor="#FFFFFF">&nbsp;<img src="../Images/small_left.gif" width="8" height="8" border=0 ../Images/small_left.gif>&nbsp;<font color=red><b><%=rs("LarCode")%></b></font>
<a alt="��<font color=red> <%=rs("LarCode")%> </font>����Ӷ�������" href='manage_product.asp?action=addMidclass&LarCode=<%=rs("LarCode")%>'>����Ӷ������ࣩ</a></td>
<td bgcolor="#FFFFFF"> <u alt="�༭ <font color=red><%=rs("LarCode")%></font> ������" style="cursor:hand" onclick="{location.href='manage_product.asp?action=editclass&LarCode=<%=rs("LarCode")%>';}">�༭</u>&nbsp;&nbsp;<a href='manage_product.asp?action=larup&larcode=<%=rs("LarCode")%>'><img src="../Images/small_up1.gif" alt='һ�������ƶ�' width="11" height="12" border=0 ../Images/small_up1.gif></a>&nbsp;&nbsp;<a href='manage_product.asp?action=lardown&larcode=<%=rs("LarCode")%>'><img src="../Images/small_down1.gif" alt='һ�����������ƶ�' width="11" height="12" border=0 ../Images/small_down1.gif></a>&nbsp;&nbsp;<u 
<%if tree_del_yesorno="1" then %>
alt="ɾ�� <font color=red><%=rs("LarCode")%></font> ����<br>ע�⣺�˲�����ͬʱɾ���÷����µ�ȫ����Ʒ<br>��ʾ�����ɵ� <font color=red>��������</font> ��ȥ��������" 
<% else %>
alt="ɾ�� <font color=red><%=rs("LarCode")%></font> ����<br>ע�⣺ɾ�������������Ʒ��תΪδ������Ʒ<br>��ʾ�����ɵ� <font color=red>��������</font> ȥ��������" 
<%end if%>
style="cursor:hand" value="ɾ��" onclick="{if(confirm('ע�⣺�ò������ɻָ���\n\nȷʵɾ����ѡһ��������')){location.href='manage_product.asp?action=delLarclass&LarCode=<%=rs("LarCode")%>';}}">ɾ��</u>   &nbsp; <a href='manage_product.asp?action=view&larcode=<%=rs("LarCode")%>' alt="���<font color=red> <%=rs("LarCode")%> </font>�����µ�ȫ����Ʒ">���</a>
<%if addlar<>"no" then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=manage_product.asp?action=addLarclass><font color=blue><b>���һ������</b></font></a>"%></td></tr>
<%
addlar="no"
end if
%>
<tr><td width=50% height=25 bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;<font color=#0066cc>
<b><%=rs("MidCode")%></b></font></td>
<td bgcolor="#FFFFFF"><u alt="�޸� <font color=red><%=rs("MidCode")%></font> ������" style="cursor:hand" onclick="{location.href='manage_product.asp?action=editclass&MidCode=<%=rs("MidCode")%>&larcode=<%=rs("LarCode")%>';}">�޸�</u>&nbsp;&nbsp;<a href='manage_product.asp?action=midup&MidCode=<%=rs("MidCode")%>&larcode=<%=rs("LarCode")%>'><img src="../Images/small_up2.gif" alt='�������������ƶ�' width="11" height="12" border=0 ../Images/small_up2.gif></a>&nbsp;&nbsp;<a href='manage_product.asp?action=middown&MidCode=<%=rs("MidCode")%>&larcode=<%=rs("LarCode")%>'><img src="../Images/small_down2.gif" alt='�������������ƶ�' width="11" height="12" border=0 ../Images/small_down2.gif></a>&nbsp;&nbsp;<u 
<%if tree_del_yesorno="1" then%>
alt="ɾ�� <font color=red><%=rs("MidCode")%></font> ����<br>ע�⣺�˲�����ͬʱɾ���÷����µ�ȫ����Ʒ<br>��ʾ�����ɵ� <font color=red>��������</font> ��ȥ��������" 
<%else%>
alt="ɾ�� <font color=red><%=rs("MidCode")%></font> ����<br>ע�⣺ɾ�������������Ʒ��תΪδ������Ʒ<br>��ʾ�����ɵ� <font color=red>��������</font> ȥ��������" 
<%end if%>
style="cursor:hand" onclick="{if(confirm('ע�⣺�ò������ɻָ���\n\nȷʵɾ����ѡ����������')){location.href='manage_product.asp?action=delMidclass&MidCode=<%=rs("MidCode")%>&larcode=<%=rs("LarCode")%>';}}">ɾ��</u> &nbsp; <a href='manage_product.asp?action=view&MidCode=<%=rs("MidCode")%>&larcode=<%=rs("LarCode")%>' alt="���<font color=red> <%=rs("MidCode")%> </font>�����µ�ȫ����Ʒ">���</a></td>
</tr>
<%
old=rs("LarCode")
rs.movenext
if rs.eof then exit do
Loop
end if
response.write "<tr><td colspan=2 bgcolor=#FFFFFF><input type='button' onClick=""javascript:location.href='manage_product.asp?action=output';"" value='���·���' alt='�޸ķ��������·���'></td></tr>"
rs.close
set rs=nothing
end sub

'����������������
if action="midup" then
LarCode=request("LarCode")
MidCode=request("MidCode")
dim i,mid(100)
i=0
set rs=server.createobject("adodb.recordset")
sql="select * from dk501_productclass where larcode='"&larcode&"' order by MidSeq asc"
rs.open sql,conn,1,3
do while not rs.eof
i=i+1
if rs("midcode")=midcode then G=i
if rs.eof then exit do
rs.movenext
loop
rs.movefirst
i=0
do while not rs.eof
i=i+1
if i=g-1 then
rs("midseq")=i+1
elseif i>1 and i=g then
rs("midseq")=i-1
else
rs("midseq")=i
end if
rs.update
if rs.eof then exit do
rs.movenext
loop
response.Redirect "manage_product.asp"
rs.close
set rs=nothing
end if

'�����������½�
if action="middown" then
LarCode=request("LarCode")
MidCode=request("MidCode")
i=0
set rs=server.createobject("adodb.recordset")
sql="select * from dk501_productclass where larcode='"&larcode&"' order by MidSeq asc"
rs.open sql,conn,1,3
do while not rs.eof
i=i+1
if rs("midcode")=midcode then G=i
if rs.eof then exit do
rs.movenext
loop
rs.movefirst
i=0
do while not rs.eof
i=i+1
if i=g then
rs("midseq")=i+1
elseif i=g+1 then
rs("midseq")=i-1
else
rs("midseq")=i
end if
rs.update
if rs.eof then exit do
rs.movenext
loop
response.Redirect "manage_product.asp"
rs.close
set rs=nothing
end if

'һ��������������
if action="larup" then
LarCode=request("LarCode")
i=0
set rs=server.createobject("adodb.recordset")
sql="select * from dk501_productclass order by LarSeq asc, MidSeq asc"
rs.open sql,conn,1,3
old=""
do while not rs.eof
if rs("larcode")<>old then i=i+1
if rs("larcode")=larcode then
G=i
exit do
end if
if rs.eof then exit do
old=rs("larcode")
rs.movenext
loop
rs.movefirst
i=0
old=""
do while not rs.eof
if rs("larcode")<>old then i=i+1
if i=g-1 then
rs("larseq")=i+1
elseif i>1 and i=g then
rs("larseq")=i-1
else
rs("larseq")=i
end if
rs.update
if rs.eof then exit do
old=rs("larcode")
rs.movenext
loop
response.Redirect "manage_product.asp"
rs.close
set rs=nothing
end if

'һ���������½�
if action="lardown" then
LarCode=request("LarCode")
i=0
set rs=server.createobject("adodb.recordset")
sql="select * from dk501_productclass order by LarSeq asc, MidSeq asc"
rs.open sql,conn,1,3
old=""
do while not rs.eof
if rs("larcode")<>old then i=i+1
if rs("larcode")=larcode then
G=i
exit do
end if
if rs.eof then exit do
old=rs("larcode")
rs.movenext
loop
rs.movefirst
i=0
old=""
do while not rs.eof
if rs("larcode")<>old then i=i+1
if i=g then
rs("larseq")=i+1
elseif i=g+1 then
rs("larseq")=i-1
else
rs("larseq")=i
end if
rs.update
if rs.eof then exit do
old=rs("larcode")
rs.movenext
loop
response.Redirect "manage_product.asp"
rs.close
set rs=nothing
end if


'���·���
if action="output" then

    Set Fso = CreateObject("Scripting.FileSystemObject")
	set dk501_productclass=Fso.CreateTextFile(Server.MapPath("../include/class.asp"),true)


	set rs=conn.execute("select * from dk501_productclass order by LarSeq asc, MidSeq asc")
	if rs.eof and rs.bof then 
		dk501_productclass.writeline "������Ʒ����"
	else
		Do While Not rs.eof
		if rs("LarCode")<>old then
			dk501_productclass.writeline "lar:"&rs("LarCode")
		end if
			dk501_productclass.writeline rs("MidCode")
			old=rs("LarCode")
		rs.movenext
		if rs.eof then exit do
		Loop
	end if
	set rs=nothing

	dk501_productclass.close
    set fso=nothing

response.write "<script language='javascript'>"
response.write "alert('�����ɹ�����Ʒ����Ѹ��£���ҳ����ʾ���º����Ʒ���ࡣ');"
response.write "location.href='manage_product.asp';"
response.write "</script>"
response.end
end if

if action="view" then
if request("midcode")<>"" then
sql="select * from dk501_product where ProdId is not null and larcode='"&request("larcode")&"' and midcode='"&request("midcode")&"'"
else
sql="select * from dk501_product where ProdId is not null and larcode='"&request("larcode")&"'"
end if
response.cookies("buyok")("sql")=sql
Response.Redirect("manage_prod1.asp")
end if
%>
		  <table height="5"><tr><td></td></tr></table>
