<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/connect.asp"-->
<%
call aspsql()
id=trim(request("id"))
if id="" then response.redirect "/index.asp"
	Set rsprod=conn.execute ("select id,message_title,message_body from dk501_message where  id='"&id&"'")
	if (rsprod.bof and rsprod.eof) then
		response.redirect "../index.asp"
	else
		message_title=rsprod("message_title")
		message_body=rsprod("message_body")
%>
<% end if %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
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
<!--#include file="include/web_top.asp" -->
<div id='div_Public'>
  <div id='Contact_top'>
    <ul>
      <li>��ǰλ�ã���ҳ > ���� > ������֪</li>
    </ul>
  </div>
  <div id='Contact_top_left'>
    <div id="sidepanel">
      <h2>���ٵ���</h2>
      <ul>
        <li>�� <a href="Contact.asp?id=3">��ͨ���</a></li>
        <li>�� <a href="Contact.asp?id=5">���ʷ��Ż�</a></li>
        <li>�� <a href="Contact.asp?id=6">��ϵ����</a></li>
        <li>�� <a href="Contact.asp?id=7">��λ���</a></li>
        <li>�� <a href="Contact.asp?id=8">������֪</a></li>
        <li>�� <a href="Contact.asp?id=9">����˿�</a></li>
      </ul>
    </div>
    <div id="sidepanel">
      <h2>�ͻ���������</h2>
      <ul>
        <li>���䣺307083095@qq.com</li>
        <li>�绰��<%=adm_tel%></li>
        <li>����������ϵ���߿ͷ�</li>
      </ul>
    </div>
  </div>
  <div id='Contact_top_right'>
    <div id="sidepanel">
      <h2><%=message_title%></h2>
      <ul><%=message_body%></ul>
    </div>
  </div>
</div>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>
