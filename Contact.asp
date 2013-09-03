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
      <li>当前位置：首页 > 裁男 > 代理须知</li>
    </ul>
  </div>
  <div id='Contact_top_left'>
    <div id="sidepanel">
      <h2>快速导航</h2>
      <ul>
        <li>◇ <a href="Contact.asp?id=3">普通快递</a></li>
        <li>◇ <a href="Contact.asp?id=5">免邮费优惠</a></li>
        <li>◇ <a href="Contact.asp?id=6">联系我们</a></li>
        <li>◇ <a href="Contact.asp?id=7">如何换货</a></li>
        <li>◇ <a href="Contact.asp?id=8">代理须知</a></li>
        <li>◇ <a href="Contact.asp?id=9">如何退款</a></li>
      </ul>
    </div>
    <div id="sidepanel">
      <h2>客户服务中心</h2>
      <ul>
        <li>邮箱：307083095@qq.com</li>
        <li>电话：<%=adm_tel%></li>
        <li>其它：请联系在线客服</li>
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
