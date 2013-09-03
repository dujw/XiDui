<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<%
dim ObjInstalled,Action,FoundErr,ErrMsg
Action=trim(request("Action"))
if Action="" then
	Action="ShowInfo"
end if
%>
 
<%
if Action="SaveConfig" then
	call SaveConfig()
	Response.Redirect "SiteConfig.asp" 
else
	call ShowConfig()
end if

call CloseConn()

sub ShowConfig()
%>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;常规广告设置</font></B></font></TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table width="99%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
  <form action=SiteConfig.asp method=post name="adv">
    <tr>
      <td width=20% height="25" align=right bgcolor="#FFFFFF"><strong>LOGO地址：（180*45）：</strong>&nbsp;</td>
      <td height="25" bgcolor="#FFFFFF"><input name="logo" type=text value="<%=index_Logo%>" size=40 maxlength=100>
        <INPUT TYPE="button" value="浏览"  onClick="window.open('../upload/manage_upload.asp?fuptype=adv&fupname=logo&frmname=logo','blank_','scrollbars=yes,resizable=no,width=650,height=450')"></td>
    </tr>
    <tr>
      <td width=20% height="25" rowspan="5" align=right bgcolor="#FFFFFF"><strong>首页左边动画图片（710*400）：</strong>&nbsp;&nbsp;</td>
      <td height="25" bgcolor="#FFFFFF"><input name="index_left_1" type=text value="<%=index_left_1%>" size=40 maxlength=100>
        <INPUT TYPE="button" value="浏览"  onClick="window.open('../upload/manage_upload.asp?fuptype=adv&fupname=index_left_1&frmname=index_left_1','blank_','scrollbars=yes,resizable=no,width=650,height=450')">
        连接：
        <input name="index_left_url_1" type=text id="index_left_url_1" value="<%=index_left_url_1%>" size=40 maxlength=100></td>
    </tr>
    <tr>
      <td height="25" bgcolor="#FFFFFF"><input name="index_left_2" type=text value="<%=index_left_2%>" size=40 maxlength=100>
        <INPUT TYPE="button" value="浏览"  onClick="window.open('../upload/manage_upload.asp?fuptype=adv&fupname=index_left_2&frmname=index_left_2','blank_','scrollbars=yes,resizable=no,width=650,height=450')">
        连接：
        <input name="index_left_url_2" type=text id="index_left_url_2" value="<%=index_left_url_2%>" size=40 maxlength=100></td>
    </tr>
    <tr>
      <td height="25" bgcolor="#FFFFFF"><input name="index_left_3" type=text value="<%=index_left_3%>" size=40 maxlength=100>
        <INPUT TYPE="button" value="浏览"  onClick="window.open('../upload/manage_upload.asp?fuptype=adv&fupname=index_left_3&frmname=index_left_3','blank_','scrollbars=yes,resizable=no,width=650,height=450')">
        连接：
        <input name="index_left_url_3" type=text id="index_left_url_3" value="<%=index_left_3%>" size=40 maxlength=100></td>
    </tr>
    <tr>
      <td height="25" bgcolor="#FFFFFF"><input name="index_left_4" type=text value="<%=index_left_4%>" size=40 maxlength=100>
        <INPUT TYPE="button" value="浏览"  onClick="window.open('../upload/manage_upload.asp?fuptype=adv&fupname=index_left_4&frmname=index_left_4','blank_','scrollbars=yes,resizable=no,width=650,height=450')">
        连接：
        <input name="index_left_url_4" type=text id="index_left_url_4" value="<%=index_left_url_4%>" size=40 maxlength=100></td>
    </tr>
    <tr>
      <td height="25" bgcolor="#FFFFFF"><input name="index_left_5" type=text value="<%=index_left_5%>" size=40 maxlength=100>
        <INPUT TYPE="button" value="浏览"  onClick="window.open('../upload/manage_upload.asp?fuptype=adv&fupname=index_left_5&frmname=index_left_5','blank_','scrollbars=yes,resizable=no,width=650,height=450')">
        连接：
        <input name="index_left_url_5" type=text id="index_left_url_5" value="<%=index_left_url_5%>" size=40 maxlength=100></td>
    </tr>
    <tr>
      <td width=20% height="25" rowspan="3" align=right bgcolor="#FFFFFF"><strong>首页右边图片（230*130）：</strong>&nbsp;&nbsp;</td>
      <td height="25" bgcolor="#FFFFFF"><input name="index_right_1" type=text value="<%=index_right_1%>" size=40 maxlength=100>
        <INPUT TYPE="button" value="浏览"  onClick="window.open('../upload/manage_upload.asp?fuptype=adv&fupname=index_right_1&frmname=index_right_1','blank_','scrollbars=yes,resizable=no,width=650,height=450')">
        连接：
        <input name="index_right_url_1" type=text id="index_right_url_1" value="<%=index_right_url_1%>" size=40 maxlength=100></td>
    </tr>
    <tr>
      <td height="25" bgcolor="#FFFFFF"><input name="index_right_2" type=text value="<%=index_right_2%>" size=40 maxlength=100>
        <INPUT TYPE="button" value="浏览"  onClick="window.open('../upload/manage_upload.asp?fuptype=adv&fupname=index_right_2&frmname=index_right_2','blank_','scrollbars=yes,resizable=no,width=650,height=450')">
        连接：
        <input name="index_right_url_2" type=text id="index_right_url_2" value="<%=index_right_url_2%>" size=40 maxlength=100></td>
    </tr>
    <tr>
      <td height="25" bgcolor="#FFFFFF"><input name="index_right_3" type=text value="<%=index_right_3%>" size=40 maxlength=100>
        <INPUT TYPE="button" value="浏览"  onClick="window.open('../upload/manage_upload.asp?fuptype=adv&fupname=index_right_3&frmname=index_right_3','blank_','scrollbars=yes,resizable=no,width=650,height=450')">
        连接：
        <input name="index_right_url_3" type=text id="index_right_url_3" value="<%=index_right_url_3%>" size=40 maxlength=100></td>
    </tr>
    <tr>
      <td width=20% height="25" align=right bgcolor="#FFFFFF"><strong>首页下面左图片（430*300）：</strong>&nbsp;&nbsp;</td>
      <td height="25" bgcolor="#FFFFFF"><input name="index_adv_1" type=text value="<%=index_adv_1%>" size=40 maxlength=100>
        <INPUT TYPE="button" value="浏览"  onClick="window.open('../upload/manage_upload.asp?fuptype=adv&fupname=index_adv_1&frmname=index_adv_1','blank_','scrollbars=yes,resizable=no,width=650,height=450')">
        连接：
        <input name="index_adv_url_1" type=text id="index_adv_url_1" value="<%=index_adv_url_1%>" size=40 maxlength=100></td>
    </tr>
    <tr>
      <td width=20% height="25" align=right bgcolor="#FFFFFF"><strong>首页下面中间图片（270*300）：</strong>&nbsp;&nbsp;</td>
      <td height="25" bgcolor="#FFFFFF"><input name="index_adv_2" type=text value="<%=index_adv_2%>" size=40 maxlength=100>
        <INPUT TYPE="button" value="浏览"  onClick="window.open('../upload/manage_upload.asp?fuptype=adv&fupname=index_adv_2&frmname=index_adv_2','blank_','scrollbars=yes,resizable=no,width=650,height=450')">
        连接：
        <input name="index_adv_url_2" type=text id="index_adv_url_2" value="<%=index_adv_url_2%>" size=40 maxlength=100></td>
    </tr>
    <tr>
      <td width=20% height="25" align=right bgcolor="#FFFFFF"><strong>首页下面右边图片（270*300）：</strong>&nbsp;&nbsp;</td>
      <td height="25" bgcolor="#FFFFFF"><input name="index_adv_3" type=text value="<%=index_adv_3%>" size=40 maxlength=100>
        <INPUT TYPE="button" value="浏览"  onClick="window.open('../upload/manage_upload.asp?fuptype=adv&fupname=index_adv_3&frmname=index_adv_3','blank_','scrollbars=yes,resizable=no,width=650,height=450')">
        连接：
        <input name="index_adv_url_3" type=text id="index_adv_url_3" value="<%=index_adv_url_3%>" size=40 maxlength=100></td>
    </tr>
    <tr>
      <td height="25" align=right bgcolor="#FFFFFF">&nbsp;</td>
      <td height="25" bgcolor="#FFFFFF">
	  <input name="Action" type="hidden" id="Action" value="SaveConfig">
	  <input name="cmdSave" type="submit"  value=" 保存设置 " >
      </td>
    </tr>
  </form>
</table>
<% end sub %>
<%
sub SaveConfig()
	dim fso,hf
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	set hf=fso.CreateTextFile(Server.mappath("../../adv/config.asp"),true)
	hf.write "<" & "%" & vbcrlf
	hf.write "Const index_Logo=" & chr(34) & trim(request("logo")) & chr(34) & "        '网站名称" & vbcrlf
	hf.write "Const index_left_1=" & chr(34) & trim(request("index_left_1")) & chr(34) & "        '动画图片" & vbcrlf
	hf.write "Const index_left_2=" & chr(34) & trim(request("index_left_2")) & chr(34) & "        '动画图片" & vbcrlf
	hf.write "Const index_left_3=" & chr(34) & trim(request("index_left_3")) & chr(34) & "        '动画图片" & vbcrlf
	hf.write "Const index_left_4=" & chr(34) & trim(request("index_left_4")) & chr(34) & "        '动画图片" & vbcrlf
	hf.write "Const index_left_5=" & chr(34) & trim(request("index_left_5")) & chr(34) & "        '动画图片" & vbcrlf
	hf.write "Const index_left_url_1=" & chr(34) & trim(request("index_left_url_1")) & chr(34) & "        '动画连接" & vbcrlf
	hf.write "Const index_left_url_2=" & chr(34) & trim(request("index_left_url_2")) & chr(34) & "        '动画连接" & vbcrlf
	hf.write "Const index_left_url_3=" & chr(34) & trim(request("index_left_url_3")) & chr(34) & "        '动画连接" & vbcrlf
	hf.write "Const index_left_url_4=" & chr(34) & trim(request("index_left_url_4")) & chr(34) & "        '动画连接" & vbcrlf
	hf.write "Const index_left_url_5=" & chr(34) & trim(request("index_left_url_5")) & chr(34) & "        '动画连接" & vbcrlf
	hf.write "Const index_right_1=" & chr(34) & trim(request("index_right_1")) & chr(34) & "        '动画图片" & vbcrlf
	hf.write "Const index_right_2=" & chr(34) & trim(request("index_right_2")) & chr(34) & "        '动画图片" & vbcrlf
	hf.write "Const index_right_3=" & chr(34) & trim(request("index_right_3")) & chr(34) & "        '动画图片" & vbcrlf
	hf.write "Const index_right_url_1=" & chr(34) & trim(request("index_right_url_1")) & chr(34) & "        '动画连接" & vbcrlf
	hf.write "Const index_right_url_2=" & chr(34) & trim(request("index_right_url_2")) & chr(34) & "        '动画连接" & vbcrlf
	hf.write "Const index_right_url_3=" & chr(34) & trim(request("index_right_url_3")) & chr(34) & "        '动画连接" & vbcrlf
	hf.write "Const index_adv_1=" & chr(34) & trim(request("index_adv_1")) & chr(34) & "        '动画图片" & vbcrlf
	hf.write "Const index_adv_2=" & chr(34) & trim(request("index_adv_2")) & chr(34) & "        '动画图片" & vbcrlf
	hf.write "Const index_adv_3=" & chr(34) & trim(request("index_adv_3")) & chr(34) & "        '动画图片" & vbcrlf
	hf.write "Const index_adv_url_1=" & chr(34) & trim(request("index_adv_url_1")) & chr(34) & "        '动画连接" & vbcrlf
	hf.write "Const index_adv_url_2=" & chr(34) & trim(request("index_adv_url_2")) & chr(34) & "        '动画连接" & vbcrlf
	hf.write "Const index_adv_url_3=" & chr(34) & trim(request("index_adv_url_3")) & chr(34) & "        '动画连接" & vbcrlf
	hf.write "%" & ">"
	hf.close
	set hf=nothing
	set fso=nothing	
end sub
%>
</BODY>
</html>
