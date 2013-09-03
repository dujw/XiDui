<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../include/connect.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="../image/web_css/login.css" rel="stylesheet" type="text/css" />
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
<!--#include file="../include/login/top.asp"-->
<div id='succes'>
  <div id='success'><img src="../image/success.gif" width="41" height="39" hspace="10" vspace="10" align="absmiddle" />恭喜您，您已经注册成功，请记住你自己的用户名</div>
  <div id='success_text'>小提醒：</div>
  <div id='success_text'>1、您可以修改你的安全设置。</div>
  <div id='success_text'>2、继续购买您的新商品。</div> 
  <div id='success_text'>3、进入"我的用户" 查看购买的商品列表。</div> 
  <div id='success_text'>4、进入"我的用户 >> 收藏的产品" 查看你收藏的产品。</div> 
  <div id='success_text'></div> 
</div>
<div id='succes'>
  <div class="index_pic">
        <h4>网站 热卖 产品 推荐</h4>
        <ul>
                  <li><a target="_blank" title="实用电脑系统ico图标" href="/icon_pic/807.shtml"><img style="background:url(http://www.aa25.cn/uploadfile/200910/12/2233499697.png);" src="/v3/images/blank.gif" width="148" height="148" />实用电脑系统ico图标</a></li>
                    <li><a target="_blank" title="一组水晶web2.0风格按钮、气泡常用图标" href="/icon_pic/798.shtml"><img style="background:url(http://www.aa25.cn/uploadfile/200907/30/111599145.jpg);" src="/v3/images/blank.gif" width="148" height="148" />一组水晶web2.0风格按钮、气泡常用图标</a></li>
					<li><a target="_blank" title="一组水晶web2.0风格按钮、气泡常用图标" href="/icon_pic/798.shtml"><img style="background:url(http://www.aa25.cn/uploadfile/200907/30/111599145.jpg);" src="/v3/images/blank.gif" width="148" height="148" />一组水晶web2.0风格按钮、气泡常用图标</a></li>
					<li><a target="_blank" title="一组水晶web2.0风格按钮、气泡常用图标" href="/icon_pic/798.shtml"><img style="background:url(http://www.aa25.cn/uploadfile/200907/30/111599145.jpg);" src="/v3/images/blank.gif" width="148" height="148" />一组水晶web2.0风格按钮、气泡常用图标</a></li>
                    <li><a target="_blank" title="一组常用的水晶风格图片，含ICO/PNG两种格式" href="/icon_pic/797.shtml"><img style="background:url(http://www.aa25.cn/uploadfile/200907/30/112466705.jpg);" src="/v3/images/blank.gif" width="148" height="148" />一组常用的水晶风格图片，含ICO/PNG两种格式</a></li>
                    <li><a target="_blank" title="2009年父亲节艺术字体" href="/photo_png/786.shtml"><img style="background:url(http://www.aa25.cn/uploadfile/200906/16/1434533413.jpg);" src="/v3/images/blank.gif" width="148" height="148" />2009年父亲节艺术字体</a></li>
        </ul>
      </div>
</div>
<!--#include file="../include/login/bootom.asp"-->
</body>
</html>
<% call CloseConn() %>