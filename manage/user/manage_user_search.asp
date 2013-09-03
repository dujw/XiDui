<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<%
action=request("action")
if action="" then
%>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;搜索会员</font></B></font>&nbsp;</TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F" width="99%">
<form name=search action=manage_user_search.asp method=post>
  <tr bgcolor="#FFFFFF">
    <td width=20% height="25" align=center>根据会员级别</td>
    <td height="25"> 
     <%	
		response.write "<select name='usertype'>"
		response.write	"<option value='' selected>选择要搜索的会员级别</option>"
		if type1<>"" then response.write "<option value='1'>"&type1&"</option>"
		if type2<>"" then response.write "<option value='2'>"&type2&"</option>"
		if type3<>"" then response.write "<option value='3'>"&type3&"</option>"
		if type4<>"" then response.write "<option value='4'>"&type4&"</option>"
		if type5<>"" then response.write "<option value='5'>"&type5&"</option>"
		if type6<>"" then response.write "<option value='6'>"&type6&"</option>"
	%>    </td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td width=20% height="25" align=center> 根据会员注册ID </td>
    <td height="25"><input type=text name=userid size=25 maxlength=20><span class="800_title">※ 会员帐号</span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
  <td width=20% height="25" align=center bgcolor="#CCCCCC">根据会员姓名 </td>
  <td height="25" bgcolor="#CCCCCC"><input type=text name=username size=25 maxlength=20><span class="800_title">※ 会员名字</span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
  <td width=20% height="25" align=center>根据会员电子邮箱 </td>
  <td height="25"><input type=text name=usermail size=25 maxlength=100></td>
  </tr>
  <tr bgcolor="#FFFFFF">
  <td width=20% height="25" align=center>根据会员联系电话 </td>
  <td height="25"><input type=text name=HomePhone size=25 maxlength=20></td></tr>
  <tr bgcolor="#FFFFFF">
  <td width=20% height="25" align=center>根据会员联系地址 </td>
  <td height="25"><input type=text name=address size=25 maxlength=100></td></tr>
  <tr bgcolor="#FFFFFF">
  <td width=20% height="25" align=center bgcolor="#CCCCCC">根据会员QQ号码 </td>
  <td height="25" bgcolor="#CCCCCC"><input name=UserQQ type=text id="UserQQ" size=25 maxlength=9><span class="800_title">※ 会员QQ号码</span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
  <td width=20% height="25" align=center bgcolor="#CCCCCC">根据会员淘宝号码 </td>
  <td height="25" bgcolor="#CCCCCC"><input name=taobao_name type=text id="taobao_name" size=25><span class="800_title">※ 会员淘宝的帐号 例：1688sc</span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
  <td width=20% height="25" align=center bgcolor="#CCCCCC">根据会员淘宝地址 </td>
  <td height="25" bgcolor="#CCCCCC"><input name=taobao type=text id="taobao" size=25><span class="800_title">※ 会员淘宝的店铺地址 例：373793071.paipai.com</span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
  <td width=20% height="25" align=center>根据会员注册时间 </td>
  <td height="25"><input type="radio" name="SignDate" value="1">
    今天 &nbsp; <input type="radio" name="SignDate" value="2">
    昨天 &nbsp; <input type="radio" name="SignDate" value="7">
    一周内 &nbsp; <input type="radio" name="SignDate" value="30">
    一月内 &nbsp; <!--<input type="radio" name="SignDate" value="">
    全部 &nbsp; --></td></tr>
	  <tr bgcolor="#FFFFFF">

  <td width=20% height="25" align=center>根据最后登陆时间 </td>
  <td height="25"><input type="radio" name="lastlogin" value="1">
    今天 &nbsp; <input type="radio" name="lastlogin" value="2">
    昨天 &nbsp; <input type="radio" name="lastlogin" value="7">
    一周内 &nbsp; <input type="radio" name="lastlogin" value="30">
    一月内 &nbsp; <!--<input type="radio" name="lastlogin" value="">
    全部 &nbsp; --></td></tr>
	  <tr bgcolor="#FFFFFF">

  <td width=20% height="25" align=center>根据会员登陆次数 </td>
  <td height="25"><input type=text name=TotalLogin1 size=4 maxlength=3>
    次至
    <input type=text name=TotalLogin2 size=4 maxlength=3>
    次</td></tr>
	  <tr bgcolor="#FFFFFF">

  <td width=20% height="25" align=center>根据会员IP上次登陆IP </td>
  <td height="25"><input type=text name=IP size=25 maxlength=15></td></tr>
    <tr bgcolor="#FFFFFF">

  <td width=20% height="25" align=center bgcolor="#CCCCCC">根据会员状态 </td>
  <td height="25" bgcolor="#CCCCCC"><input type="radio" name="Status" value="1">
    正常 
    <input type="radio" name="Status" value="0">
    锁定/审核中</td></tr>
  <tr bgcolor="#FFFFFF">
    <td height="25" colspan=2><INPUT name="action" TYPE="hidden" value="ok">
      <INPUT name=action TYPE="submit" value="开始搜索"></td>
  </tr>
  </form>
</table>
<%
else
	if trim(request("TotalLogin1"))<>"" then
	  if not isNumeric(trim(request("TotalLogin1"))) then
		response.write "<script language='javascript'>"
		response.write "alert('出错了，登陆次数必须是数字！');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	  end if
	end if
	if trim(request("TotalLogin2"))<>"" then
	  if not isNumeric(trim(request("TotalLogin2"))) then
		response.write "<script language='javascript'>"
		response.write "alert('出错了，登陆次数必须是数字！');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	  end if
	end if
	usersql=""
	if request("usertype")<>"" then usersql=usersql+" and UserType='"&request("usertype")&"'"
	if trim(request("userID"))<>"" then usersql=usersql+" and userid like '%"&trim(request("userid"))&"%'"
	if trim(request("username"))<>"" then usersql=usersql+" and username like '%"&trim(request("username"))&"%'"
	if trim(request("userqq"))<>"" then usersql=usersql+" and userqq like '%"&trim(request("userqq"))&"%'"
	
	if trim(request("taobao_name"))<>"" then usersql=usersql+" and taobao_name like '%"&trim(request("taobao_name"))&"%'"
	if trim(request("taobao"))<>"" then usersql=usersql+" and taobao like '%"&trim(request("taobao"))&"%'"
	
	if trim(request("usermail"))<>"" then usersql=usersql+" and usermail like '%"&trim(request("usermail"))&"%'"
	if trim(request("address"))<>"" then usersql=usersql+" and address like '%"&trim(request("address"))&"%'"
	if trim(request("HomePhone"))<>"" then usersql=usersql+" and HomePhone like '%"&trim(request("HomePhone"))&"%'"
	if trim(request("ip"))<>"" then usersql=usersql+" and ip like '%"&trim(request("ip"))&"%'"
	if request("SignDate")="1" then  usersql=usersql+" and datediff(day,[SignDate],getdate())=0"
	if request("SignDate")="2" then  usersql=usersql+" and datediff(day,[SignDate],getdate())=1"
	if request("SignDate")="7" then  usersql=usersql+" and datediff(week,[SignDate],getdate())=0"
	if request("SignDate")="30" then  usersql=usersql+" and datediff(month,[SignDate],getdate())=0"
	if request("lastlogin")="1" then  usersql=usersql+" and datediff(day,[lastlogin],getdate())=0"
	if request("lastlogin")="2" then  usersql=usersql+" and datediff(day,[lastlogin],getdate())=1"
	if request("lastlogin")="7" then  usersql=usersql+" and datediff(week,[lastlogin],getdate())=0"
	if request("lastlogin")="30" then  usersql=usersql+" and datediff(month,[lastlogin],getdate())=0"

	if trim(request("TotalLogin1"))<>"" then  usersql=usersql+" and totallogin >= "& trim(request("TotalLogin1"))
	if trim(request("TotalLogin2"))<>"" then  usersql=usersql+" and totallogin <= "& trim(request("TotalLogin2"))
	if request("Status")<>"" then usersql=usersql+" and Status='"&request("Status")&"'"

	response.cookies("buyok")("usersql")=usersql
	if usersql="" then
		Response.Redirect("manage_user_search.asp")
	else
		Response.Redirect("manage_user_list.asp")
	end if
end if
%>
<br>
</BODY>
</html>
<% call CloseConn() %>
