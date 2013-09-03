<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/connect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="../image/web_css/login.css" rel="stylesheet" type="text/css" />
<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
<style type="text/css">
#succes {width:950px; margin:0 auto; border:1px solid #CCCCCC; margin-top:10px;}
#success{padding-left:80px;height:60px;color:#FF4E00;text-decoration:none;font-size: 14px;font-weight: bold;}
#success_text{ padding-left:90px;height:25px;}
.index_pic { }
.index_pic h4 { height:23px; padding:5px 8px 0 8px; color:#5f6e60;border-bottom:1px solid #dadada; font-size:14px;}
.index_pic ul { height:165px;background:#fff;padding:6px 0 8px 0; }
.index_pic ul li { padding:4px 4px 4px 4px;width:150px; height:160px; overflow:hidden; float:left;  display:inline;}
.index_pic ul li img { display:block; border:1px solid #ddd;}
.index_pic ul li a {
	display:block;
	text-align:center;
	color:#000000;
	text-decoration: none;
}
.index_pic ul li a:hover { text-decoration:underline;}
.index_pic ul li a:hover img { border:1px solid #666;}
</style>
</head>
<body> 
<!--#include file="include/web_top.asp" -->
<div id='succes'>
  <div id='success'><img src="image/success.gif" width="41" height="39" hspace="10" vspace="10" align="absmiddle" />恭喜您，您已经注册成功，请记住你自己的用户名</div>
  <div id='success_text'>小提醒：</div>
  <div id='success_text'>1、您可以修改你的安全设置。</div>
  <div id='success_text'>2、继续购买您的新商品。</div> 
  <div id='success_text'>3、进入"我的用户" 查看购买的商品列表。</div> 
  <div id='success_text'>4、进入"我的用户 >> 收藏的产品" 查看你收藏的产品。</div> 
  <div id='success_text'></div> 
</div>

<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>