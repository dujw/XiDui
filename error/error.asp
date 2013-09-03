<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../include/connect.asp" -->
<%error=request("error")%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="../image/web_css/login.css" rel="stylesheet" type="text/css" />
<link href="../image/web_css/index.css" rel="stylesheet" type="text/css" />
<style type="text/css">
#succes {width:950px; margin:0 auto; border:1px solid #CCCCCC; margin-top:10px;}
#success{padding:80px 0 80px 80px;;color:#FF4E00;text-decoration:none;font-size: 14px;font-weight: bold;}
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
<!--#include file="../include/web_top.asp" -->
<div id='succes'>
  <div id='success'>
  <%
if error="001" then response.write "提示：您不是从本服务器上提交的表单信息！ <br><a href=../index/main.asp><b>返回首页</b></a>"
if error="002" then response.write "提示：您所选择的用户名已经被其他人注册了<br><br>请检查后重新 <a href=../reg_member.asp><b>注册</b></a> 新的帐号"
if error="003" then response.write "提示：没有您所填写的用户名，请检查后重新 <a href=../main.asp><b>登陆</b></a> 或者 <a href=../reg_member.asp><b>注册</b></a> 新的帐号"
if error="0031" then response.write "提示：对不起，您输入的会员号已被锁定或未通过审核。<br>请检查后重新 <a href=../main.asp><b>登陆</b></a> 或者 <a href=../reg_member.asp><b>注册</b></a> 新的帐号"
if error="004" then response.write "提示：您输入的密码不正确，请检查后重新 <a href=../main.asp><b>登陆</b></a> 或者 <a href=../reg_member.asp><b>注册</b></a> 新的帐号"
if error="005" then response.write "提示：您的帐号已经登陆，请不要重复登陆，请<a href=../main.asp><b>返回首页</b></a>"
if error="006" then response.write "提示：您没有注册或者没有登陆，请返回首页登陆或注册后再打开购物车<br>请放心，您已经放入购物车的商品不会丢失<br><br>请您 <A href=../main.asp><b>返回首页</b></a> 后登陆或注册"
if error="007" then response.write "提示：您的购物车中没有任何商品，请返回首页重新购买<br><br>请您 <A href=../index.asp><b>返回首页</b></a> 后重新购物，或者去 <a href=../fav.asp><b>收藏箱</b></a> 去看看"
if error="008" then response.write "提示：没有此用户名，或者该用户名已被管理员锁定<br><br>请检查后重新 <a href=../main.asp><b>登陆</b></a> 或者 <a href=../reg_member.asp><b>注册</b></a> 新的帐号"
if error="009" then response.write "提示：只有注册为会员并登陆后，才能使用收藏功能。<br><br>请您 <A href=../main.asp><b>返回首页</b></a> 后登陆或注册"
if error="010" then response.write "提示：您的收藏箱中没有存放任何宝贝。<br><br>您可 <A href=../main.asp><b>到首页</b></a> 去逛逛，或者去 <A href=../'shop.asp'><b>购物车</b></a> 去看看"
if error="011" then response.write "提示：无法打开您收藏的商品，可能是该商品已经下架。<br><br>请您 <A href=../main.asp><b>返回首页</b></a> 去逛逛吧"
if error="012" then response.write "提示：对不起，您的收藏箱已经装满，请您清理之后再放入新的商品。<br><br> <A href=../'fav.asp'><b>现在就去整理收藏箱</b></a>"
if error="99" then response.write "提示：您的IP地址<font color=red> "&Request.serverVariables("REMOTE_ADDR")&" </font>被本站限制访问，请与管理员联系。<br><br><img src=images/small/e-mail.gif border=0 align=center> <a href=../'mailto:"&sitename &"<"& adm_mail &">"&"?subject=IP被锁定'>"&adm_mail&"</a>"
%>
  </div>
</div>

<!--#include file="../include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>