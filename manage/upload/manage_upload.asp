<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<%
session("fuptype")=request("fuptype")		
session("fupname")=request("fupname")		
session("frmname")=request("frmname")	
response.Write session("fuptype")
response.Write session("fupname")
response.Write session("frmname")
Server.ScriptTimeOut=99999
%>
<html>
<head>
<title>�ļ��ϴ�</title>
<meta name="Description" Content="">
<link href="../images/css/admin_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#D9EAFC">
<table align="left"><tr><td>
<form name="form1" method="post" action="manage_upsave.asp" enctype="multipart/form-data">
<b>��ѡ��Ҫ�ϴ����ļ���</b><br>
<input type=file name="file1">
<input type=submit name="submit" value="�ϴ�"><br><br>
��������ϴ����������ĵȴ�����Ҫ�ظ�������ϴ��������ϴ�ʱ�����ļ���С������״������<br>
��Ϊ��ʡ�ռ䣬�����ͼƬ�ļ����뾡���Ż������鵥���ļ���Ҫ����50KB��<br>
�����ʹ��ļ�ʱ�����ܵ��·������������߲��ȶ�������ʹ��FTP�ϴ����ļ���
</form>
</td></tr></table> 
</body>  
</html>
