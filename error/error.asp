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
if error="001" then response.write "��ʾ�������Ǵӱ����������ύ�ı���Ϣ�� <br><a href=../index/main.asp><b>������ҳ</b></a>"
if error="002" then response.write "��ʾ������ѡ����û����Ѿ���������ע����<br><br>��������� <a href=../reg_member.asp><b>ע��</b></a> �µ��ʺ�"
if error="003" then response.write "��ʾ��û��������д���û�������������� <a href=../main.asp><b>��½</b></a> ���� <a href=../reg_member.asp><b>ע��</b></a> �µ��ʺ�"
if error="0031" then response.write "��ʾ���Բ���������Ļ�Ա���ѱ�������δͨ����ˡ�<br>��������� <a href=../main.asp><b>��½</b></a> ���� <a href=../reg_member.asp><b>ע��</b></a> �µ��ʺ�"
if error="004" then response.write "��ʾ������������벻��ȷ����������� <a href=../main.asp><b>��½</b></a> ���� <a href=../reg_member.asp><b>ע��</b></a> �µ��ʺ�"
if error="005" then response.write "��ʾ�������ʺ��Ѿ���½���벻Ҫ�ظ���½����<a href=../main.asp><b>������ҳ</b></a>"
if error="006" then response.write "��ʾ����û��ע�����û�е�½���뷵����ҳ��½��ע����ٴ򿪹��ﳵ<br>����ģ����Ѿ����빺�ﳵ����Ʒ���ᶪʧ<br><br>���� <A href=../main.asp><b>������ҳ</b></a> ���½��ע��"
if error="007" then response.write "��ʾ�����Ĺ��ﳵ��û���κ���Ʒ���뷵����ҳ���¹���<br><br>���� <A href=../index.asp><b>������ҳ</b></a> �����¹������ȥ <a href=../fav.asp><b>�ղ���</b></a> ȥ����"
if error="008" then response.write "��ʾ��û�д��û��������߸��û����ѱ�����Ա����<br><br>��������� <a href=../main.asp><b>��½</b></a> ���� <a href=../reg_member.asp><b>ע��</b></a> �µ��ʺ�"
if error="009" then response.write "��ʾ��ֻ��ע��Ϊ��Ա����½�󣬲���ʹ���ղع��ܡ�<br><br>���� <A href=../main.asp><b>������ҳ</b></a> ���½��ע��"
if error="010" then response.write "��ʾ�������ղ�����û�д���κα�����<br><br>���� <A href=../main.asp><b>����ҳ</b></a> ȥ��䣬����ȥ <A href=../'shop.asp'><b>���ﳵ</b></a> ȥ����"
if error="011" then response.write "��ʾ���޷������ղص���Ʒ�������Ǹ���Ʒ�Ѿ��¼ܡ�<br><br>���� <A href=../main.asp><b>������ҳ</b></a> ȥ����"
if error="012" then response.write "��ʾ���Բ��������ղ����Ѿ�װ������������֮���ٷ����µ���Ʒ��<br><br> <A href=../'fav.asp'><b>���ھ�ȥ�����ղ���</b></a>"
if error="99" then response.write "��ʾ������IP��ַ<font color=red> "&Request.serverVariables("REMOTE_ADDR")&" </font>����վ���Ʒ��ʣ��������Ա��ϵ��<br><br><img src=images/small/e-mail.gif border=0 align=center> <a href=../'mailto:"&sitename &"<"& adm_mail &">"&"?subject=IP������'>"&adm_mail&"</a>"
%>
  </div>
</div>

<!--#include file="../include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>