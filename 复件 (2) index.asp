<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/connect.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="description" content="<%=sitedescription%>">
<meta name="keywords" content="<%=sitekeywords%>">
<link href="images/css/index.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div id="div_Public">
  <div id="index_top_left"><img src="adv/logo/logo.jpg" /></div>
  <div id="index_top_right">[���¼] [���ע��] �û����� ��Ҫ���� ��Ա���� ���߿ͷ� ��������</div>
</div>
<div id="div_Public">
  <div id="index_banner">
    <ul>
      <li><a href=index.asp>��ҳ</a></li>
      <li><a href=index.asp>��Ҫ����</a></li>
      <li><a href=index.asp>��Ҫ����</a></li>
      <li><a href=index.asp>Ʒ��</a></li>
      <li><a href=index.asp>��װ</a></li>
      <li><a href=index.asp>Ůװ</a></li>
      <li><a href=index.asp>�˶�</a></li>
      <li><a href=index.asp>��Ʒ</a></li>
      <li><a href=index.asp>���</a></li>
	  <li><a href=index.asp>��ѯ</a></li>
    </ul>
  </div>
  <div id="index_banner_dowm">
    <SELECT id=type name=type> <OPTION value=1 selected>������Ʒ</OPTION></SELECT>
    &nbsp;<input name="name_key" type="text" id="name_key"  style="width:280px;" value="<%=trim(request("name_key"))%>" maxlength="140" />
	<%
	set rs_1=server.createobject("adodb.recordset") 
	sql="select distinct LarCode,LarSeq from dk501_productclass order by LarSeq" 
	rs_1.Open sql,conn,1,1 
		response.write "<select name='fenlei' class=5354_inputcolor>"
		response.write	"<option value='' selected>�������</option>"
	if rs_1.bof and rs_1.eof then
		response.write "<option selected value=''>�������</option>"
	else
		Do While Not rs_1.eof	
		response.write "<option value='"&rs_1("larCode")&"'>"&rs_1("LarCode")&"</option>"
		rs_1.movenext
		  if rs_1.eof then exit do
		loop
		response.write "<option value='δ����'>����</option>"
	end if
	rs_1.close
	set rs_1=nothing
	%>
     &nbsp;<input type="submit" name="button32" id="button32" value="����" class="search-input"/>
  </div>
</div>
</body>
</html>
<% call CloseConn() %>
