<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<table align="left">
<tr>
<td>
<%
fuptype=session("fuptype")
fupname=session("fupname")
frmname=session("frmname")
if fuptype="" or fupname="" or frmname="" then
response.write "<script language='javascript'>"
response.write "alert('���ִ����������ϴ���');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
%>
<!--#include FILE="upload_5xsoft.inc"-->
<%
set upload=new upload_5xsoft
set file=upload.file("file1")
if file.fileSize>0 then
filename=fupname+"."
filenameend=file.filename
filenameend=split(filenameend,".")
n=UBound(filenameend)
filename=filename&filenameend(n)
if fuptype<>"db" then
if file.fileSize>10000000 then
response.write "<script language='javascript'>"
response.write "alert('���ϴ����ļ�̫���ϴ����ɹ��������ļ�����ܳ���1M��');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
end if
if LCase(filenameend(n))<>"gif" and LCase(filenameend(n))<>"jpg" and LCase(filenameend(n))<>"swf" then
response.write "<script language='javascript'>"
response.write "alert('�������ϴ���ѡ����ļ���ʽ������������ϴ���');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
if fuptype="adv" then
savepath="../../adv/"&filename
elseif fuptype="pic" then
savepath="../../UploadFiles/2011"&right("00"&month(date),2)&"/"&filename
elseif fuptype="pinpai" then
savepath="../pic/pinpai/"&filename
elseif fuptype="pic1" then
savepath="../pic/digi1/"&filename
elseif fuptype="link" then
savepath="../images/links/"&filename
elseif fuptype="db" then
savepath="./"&filename
end if
file.saveAs Server.mappath(savepath)
response.write "�ļ��ϴ��ɹ�!�ϴ��ļ�������·��Ϊ��"
response.write "<font color=red>"&Server.mappath(savepath)&"</font><br><br>"
response.write "<a href='"&savepath&"'  target='_blank'>���Ԥ���ϴ����ļ�</a>"
response.write "<br><br><INPUT onclick='javascript:window.close();' type=submit value='�ϴ����'>"
%>
<script language=JavaScript>
function inserttext(){
<%
if fuptype="adv" then
response.write "self.opener.adv."&frmname&".value ='adv/"&filename&"';"
elseif fuptype="pic" then
response.write "self.opener.prodtable."&frmname&".value ='UploadFiles/2011"&right("00"&month(date),2)&"/"&filename&"';"
elseif fuptype="pinpai" then
response.write "self.opener.add."&frmname&".value ='pic/pinpai/"&filename&"';"
elseif fuptype="pic1" then
response.write "self.opener.prodtable."&frmname&".value ='pic/digi1/"&filename&"';"
elseif fuptype="link" then
response.write "self.opener.prodtable."&frmname&".value ='images/links/"&filename&"';"
end if
%>
}
inserttext();
</script>
<%
else
response.write "<script language='javascript'>"
response.write "alert('�ļ����ݲ���Ϊ�գ�������ȷ����������һҳ��');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
set upload=nothing
session("fuptype")=""
session("fupname")=""
session("frmname")=""
%>
</td></tr></table>
</body> 
</html>