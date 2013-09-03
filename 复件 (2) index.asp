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
  <div id="index_top_right">[请登录] [免费注册] 用户中心 我要加盟 会员服务 在线客服 帮助中心</div>
</div>
<div id="div_Public">
  <div id="index_banner">
    <ul>
      <li><a href=index.asp>首页</a></li>
      <li><a href=index.asp>我要进货</a></li>
      <li><a href=index.asp>我要供货</a></li>
      <li><a href=index.asp>品牌</a></li>
      <li><a href=index.asp>男装</a></li>
      <li><a href=index.asp>女装</a></li>
      <li><a href=index.asp>运动</a></li>
      <li><a href=index.asp>饰品</a></li>
      <li><a href=index.asp>配件</a></li>
	  <li><a href=index.asp>咨询</a></li>
    </ul>
  </div>
  <div id="index_banner_dowm">
    <SELECT id=type name=type> <OPTION value=1 selected>搜索商品</OPTION></SELECT>
    &nbsp;<input name="name_key" type="text" id="name_key"  style="width:280px;" value="<%=trim(request("name_key"))%>" maxlength="140" />
	<%
	set rs_1=server.createobject("adodb.recordset") 
	sql="select distinct LarCode,LarSeq from dk501_productclass order by LarSeq" 
	rs_1.Open sql,conn,1,1 
		response.write "<select name='fenlei' class=5354_inputcolor>"
		response.write	"<option value='' selected>所有类别</option>"
	if rs_1.bof and rs_1.eof then
		response.write "<option selected value=''>暂无类别</option>"
	else
		Do While Not rs_1.eof	
		response.write "<option value='"&rs_1("larCode")&"'>"&rs_1("LarCode")&"</option>"
		rs_1.movenext
		  if rs_1.eof then exit do
		loop
		response.write "<option value='未归类'>其它</option>"
	end if
	rs_1.close
	set rs_1=nothing
	%>
     &nbsp;<input type="submit" name="button32" id="button32" value="搜索" class="search-input"/>
  </div>
</div>
</body>
</html>
<% call CloseConn() %>
