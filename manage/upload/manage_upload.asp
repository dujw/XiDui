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
<title>文件上传</title>
<meta name="Description" Content="">
<link href="../images/css/admin_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#D9EAFC">
<table align="left"><tr><td>
<form name="form1" method="post" action="manage_upsave.asp" enctype="multipart/form-data">
<b>请选择要上传的文件：</b><br>
<input type=file name="file1">
<input type=submit name="submit" value="上传"><br><br>
·点击“上传”后，请耐心等待（不要重复点击“上传”），上传时间视文件大小和网络状况而定<br>
·为节省空间，如果是图片文件，请尽量优化，建议单个文件不要超过50KB。<br>
·传送大文件时，可能导致服务器变慢或者不稳定。建议使用FTP上传大文件。
</form>
</td></tr></table> 
</body>  
</html>
