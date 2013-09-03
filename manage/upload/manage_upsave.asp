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
response.write "alert('出现错误，请重新上传！');"
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
response.write "alert('您上传的文件太大，上传不成功，单个文件最大不能超过1M！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
end if
if LCase(filenameend(n))<>"gif" and LCase(filenameend(n))<>"jpg" and LCase(filenameend(n))<>"swf" then
response.write "<script language='javascript'>"
response.write "alert('不允许上传您选择的文件格式，请检查后重新上传！');"
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
response.write "文件上传成功!上传文件的物理路径为："
response.write "<font color=red>"&Server.mappath(savepath)&"</font><br><br>"
response.write "<a href='"&savepath&"'  target='_blank'>点击预览上传的文件</a>"
response.write "<br><br><INPUT onclick='javascript:window.close();' type=submit value='上传完成'>"
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
response.write "alert('文件内容不能为空，单击“确定”返回上一页！');"
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