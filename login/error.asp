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
  <div id='success'><img src="../image/success.gif" width="41" height="39" hspace="10" vspace="10" align="absmiddle" />��ϲ�������Ѿ�ע��ɹ������ס���Լ����û���</div>
  <div id='success_text'>С���ѣ�</div>
  <div id='success_text'>1���������޸���İ�ȫ���á�</div>
  <div id='success_text'>2������������������Ʒ��</div> 
  <div id='success_text'>3������"�ҵ��û�" �鿴�������Ʒ�б���</div> 
  <div id='success_text'>4������"�ҵ��û� >> �ղصĲ�Ʒ" �鿴���ղصĲ�Ʒ��</div> 
  <div id='success_text'></div> 
</div>
<div id='succes'>
  <div class="index_pic">
        <h4>��վ ���� ��Ʒ �Ƽ�</h4>
        <ul>
                  <li><a target="_blank" title="ʵ�õ���ϵͳicoͼ��" href="/icon_pic/807.shtml"><img style="background:url(http://www.aa25.cn/uploadfile/200910/12/2233499697.png);" src="/v3/images/blank.gif" width="148" height="148" />ʵ�õ���ϵͳicoͼ��</a></li>
                    <li><a target="_blank" title="һ��ˮ��web2.0���ť�����ݳ���ͼ��" href="/icon_pic/798.shtml"><img style="background:url(http://www.aa25.cn/uploadfile/200907/30/111599145.jpg);" src="/v3/images/blank.gif" width="148" height="148" />һ��ˮ��web2.0���ť�����ݳ���ͼ��</a></li>
					<li><a target="_blank" title="һ��ˮ��web2.0���ť�����ݳ���ͼ��" href="/icon_pic/798.shtml"><img style="background:url(http://www.aa25.cn/uploadfile/200907/30/111599145.jpg);" src="/v3/images/blank.gif" width="148" height="148" />һ��ˮ��web2.0���ť�����ݳ���ͼ��</a></li>
					<li><a target="_blank" title="һ��ˮ��web2.0���ť�����ݳ���ͼ��" href="/icon_pic/798.shtml"><img style="background:url(http://www.aa25.cn/uploadfile/200907/30/111599145.jpg);" src="/v3/images/blank.gif" width="148" height="148" />һ��ˮ��web2.0���ť�����ݳ���ͼ��</a></li>
                    <li><a target="_blank" title="һ�鳣�õ�ˮ�����ͼƬ����ICO/PNG���ָ�ʽ" href="/icon_pic/797.shtml"><img style="background:url(http://www.aa25.cn/uploadfile/200907/30/112466705.jpg);" src="/v3/images/blank.gif" width="148" height="148" />һ�鳣�õ�ˮ�����ͼƬ����ICO/PNG���ָ�ʽ</a></li>
                    <li><a target="_blank" title="2009�길�׽���������" href="/photo_png/786.shtml"><img style="background:url(http://www.aa25.cn/uploadfile/200906/16/1434533413.jpg);" src="/v3/images/blank.gif" width="148" height="148" />2009�길�׽���������</a></li>
        </ul>
      </div>
</div>
<!--#include file="../include/login/bootom.asp"-->
</body>
</html>
<% call CloseConn() %>