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
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;��Ʒ�߼�����</font></B></font></TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F" width="99%">
  <form name=search action=manage_product2.asp method=post>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>������Ʒ���</td>
      <td height="25"><%	Dim a,b,c 
	set rs=server.createobject("adodb.recordset") 
	sql="select distinct LarCode,LarSeq from dk501_productclass order by LarSeq" 
	rs.Open sql,conn,1,1 
	response.write "<select name='fenlei'>"
	response.write	"<option value='' selected>ѡ��Ҫ��������Ʒ����</option>"
	if rs.bof and rs.eof then
	response.write "<option selected value=''>-���޷���-</option>"
	else
	Do While Not rs.eof	
	response.write "<option value='"&rs("larCode")&"'>"&rs("LarCode")&"</option>"
	rs.movenext
	  if rs.eof then exit do
	loop
	response.write "<option value='δ����'>����δ������Ʒ</option>"
	end if
	rs.close
	set rs=nothing
%>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center> ������Ʒ���� </td>
      <td height="25"><input type=text name=name_key size=40 maxlength=20></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>������Ʒ��� </td>
      <td height="25"><input type=text name=jj_key size=40 maxlength=20></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>������Ʒ���� </td>
      <td height="25"><input type=text name=nr_key size=40 maxlength=20></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>������Ʒ�۸� </td>
      <td height="25">�۸�Χ����
        <input type=text name=jg1_key size=5 maxlength=4>
        Ԫ ��
        <input type=text name=jg2_key size=5 maxlength=4>
        Ԫ</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>�����Ƿ��Ƽ� </td>
      <td height="25"><input type="radio" name="tj_key" value="1">
        ��
        <input type="radio" name="tj_key" value="2">
        ��
        <input type="radio" name="tj_key" value="3">
        ȫ��</td>
    </TR>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>�����Ƿ��ؼ� </td>
      <td height="25"><input type="radio" name="tej_key" value="1">
        ��
        <input type="radio" name="tej_key" value="2">
        ��
        <input type="radio" name="tej_key" value="3">
        ȫ��</td>
    </TR>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>������Ʒ״̬ </td>
      <td height="25"><input type="radio" name="zt_key" value="1">
        ��ʾ
        <input type="radio" name="zt_key" value="2">
        �ر�
        <input type="radio" name="zt_key" value="3">
        ȫ�� </td>
    </TR>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>�����������ʽ </td>
      <td height="25"><input type="radio" name="paixu" value="1" checked>
        ʱ�併
        <input type="radio" name="paixu" value="2">
        ʱ����
        <input type="radio" name="paixu" value="3">
        ��Ž�
        <input type="radio" name="paixu" value="4">
        �����
        <input type="radio" name="paixu" value="5">
        ����
        <input type="radio" name="paixu" value="6">
        ����� </td>
    </TR>
    <tr bgcolor="#FFFFFF">
      <td colspan=2><INPUT name="action" TYPE="hidden" value="ok">
        <INPUT name=action TYPE="submit" value="��ʼ����"></td>
    </tr>
  </form>
</table>
<%
else

'��ʼ����
if trim(request("jg1_key"))<>"" then
  if not isNumeric(trim(request("jg1_key"))) then
	response.write "<script language='javascript'>"
	response.write "alert('����д�����ݲ�����Ҫ�󣬼۸���������֣����飡');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
  end if
end if

if trim(request("jg2_key"))<>"" then
  if not isNumeric(trim(request("jg2_key"))) then
	response.write "<script language='javascript'>"
	response.write "alert('����д�����ݲ�����Ҫ�󣬼۸���������֣����飡');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
  end if
end if

	if request("name_key")<>"" then addsql=addsql+" and prodname like '%"&request("name_key")&"%'"
	if request("jj_key")<>"" then addsql=addsql+" and ProdDisc like '%"&request("jj_key")&"%'"
	if request("nr_key")<>"" then addsql=addsql+" and MemoSpec like '%"&request("nr_key")&"%'"
	if request("tj_key")="1" then addsql=addsql+" and remark='1'"
	if request("tj_key")="2" then addsql=addsql+" and remark='0'"
	if request("zt_key")="1" then addsql=addsql+" and online=true"
	if request("zt_key")="2" then addsql=addsql+" and online=false"
	if request("tej_key")="1" then addsql=addsql+" and tejia='1'"
	if request("tej_key")="2" then addsql=addsql+" and tejia='0'"
	if request("fenlei")<>"" then addsql=addsql+" and larcode='"&request.form("fenlei")&"'"
	if trim(request("jg1_key"))<>"" then addsql=addsql+" and PriceList>"&request("jg1_key")
	if trim(request("jg2_key"))<>"" then addsql=addsql+" and PriceList<"&request("jg2_key")
	if request("paixu")="1" then orderby=" order by adddate desc"
	if request("paixu")="2" then orderby=" order by adddate"
	if request("paixu")="3" then orderby=" order by prodid desc"
	if request("paixu")="4" then orderby=" order by prodid"
	if request("paixu")="5" then orderby=" order by prodname"
	if request("paixu")="6" then orderby=" order by ClickTimes desc"

if addsql="" then
response.cookies("buyok")("sql")=""
Response.Redirect("manage_product2.asp")
else

sql="select * from bproduc where dk501_product is not null" + addsql +orderby
response.cookies("buyok")("sql")=sql
Response.Redirect("manage_prod1.asp")
end if

end if
%>
</BODY>
</html>
