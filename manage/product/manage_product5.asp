<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<%
action=request("action")
if action="" then
%>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;��Ʒ���ת��</font></B></font></TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F" width="99%">
  <form name=search action=manage_product5.asp method=post>
    <tr bgcolor="#FFFFFF">
      <td colspan=2 class=td height=25>&nbsp;��Ʒ���ת��</td>
    </tr>
    <%	Dim a,b,c 
	set rs=server.createobject("adodb.recordset") 
	sql="select distinct LarCode,LarSeq,MidCode,midSeq from dk501_productclass order by LarSeq,MIdSeq" 
	rs.Open sql,conn,1,1 
	if rs.bof and rs.eof then
	response.write "<TR bgcolor=#FFFFFF><TD align=center colspan=2>û�ж���Ʒ���з��࣬����ת����Ʒ��</td>"
	else
	response.write "<TR bgcolor=#FFFFFF><TD align=center width=100 >��</td><td>"
	response.write "<select name='fenlei1'>"
	response.write	"<option value='' selected>ѡ��Ҫת�Ƶ���Ʒ����</option>"
	Do While Not rs.eof
	a=rs("larcode")	
	response.write "<option value='"&rs("larCode")&"^"&rs("MidCode")&"'>"&rs("LarCode")&" >> "&rs("MidCode")&"</option>"
	rs.movenext
	if rs.eof then exit do
	if rs("larcode")<>a then response.write "<option value=''>--------------------------</option>"
	loop
	response.write "<option value=''>--------------------------</option>"
	response.write "<option value='δ����^δ����'>����δ�������Ʒ</option>"

	response.write "</td><tr bgcolor=#FFFFFF><TD align=center width=100 >ת�Ƶ�</td><td>"

	rs.movefirst	'ָ���ض�λ
	response.write "<select name='fenlei2'>"
	response.write	"<option value='' selected>ѡ����Ʒת�Ƶ�Ŀ��λ��</option>"
	Do While Not rs.eof
	a=rs("larcode")	
	response.write "<option value='"&rs("larCode")&"^"&rs("MidCode")&"'>"&rs("LarCode")&" >> "&rs("MidCode")&"</option>"
	rs.movenext
	if rs.eof then exit do
	if rs("larcode")<>a then response.write "<option value=''>--------------------------</option>"
	loop
	end if
	rs.close
	set rs=nothing
	response.write "</td></tr>"

%>
    <tr bgcolor="#FFFFFF">
      <td height="25" colspan=2><INPUT name=action TYPE="submit" value="��ʼת��">
        <INPUT name="action" TYPE="hidden" value="ok">
      </td>
    </tr>
  </form>
</table>
<%
else

'��ʼ����
if request("fenlei1")="" or  request("fenlei2")="" then
	response.write "<script language='javascript'>"
	response.write "alert('��û��ѡ�񣬻���û�пɹ�ѡ�����Ʒ���ࡣ');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if

if request("fenlei1")=request("fenlei2") then
	response.write "<script language='javascript'>"
	response.write "alert('ѡ����������ͬ�ķ��࣬��Ʒδת�ơ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if


code1=split(request.form("fenlei1"),"^")
larcode1=code1(0)
MidCode1=code1(1)

code2=split(request.form("fenlei2"),"^")
larcode2=code2(0)
MidCode2=code2(1)

set rs=server.createobject("adodb.recordset")
sql = "select * from dk501_product where larcode='"&larcode1&"' and midcode='"&midcode1&"'"
rs.open sql,conn,1,3
if  (rs.eof and rs.bof) then

	response.write "<script language='javascript'>"
	response.write "alert('���������û����Ʒ�ɹ�ת�ơ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end

else
i=0
do while not rs.eof
rs("larcode")=larcode2
rs("midcode")=midcode2
i=i+1
if rs.eof then exit do
rs.movenext
loop

	response.write "<script language='javascript'>"
	response.write "alert('��Ʒת�Ƴɹ���');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end

end if
rs.close
set rs=nothing


end if
%>
</BODY>
</html>
<% call CloseConn() %>
