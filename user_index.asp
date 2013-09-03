<!--#include file="include/connect.asp" -->
<%
call aspsql()
'判断用户是否登陆
call checklogin()
Name   = request.cookies("buyok")("userid")

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
</head>
<body>
<!--#include file="include/web_top.asp" -->
<div id="div_Public">
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="260" align="left" valign="top"><!--#include file="left.asp"--></td>
    <td valign="top">
	          <table width="690" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" bgcolor="E43A23">&nbsp;<span class="input_search">用户中心</span></td>
                </tr>
                <tr>
                  <td><%
Set rstxtt = conn.Execute("select * from dk501_bconfig") 
	usertype1=rstxtt("usertype1")
	usertype2=rstxtt("usertype2")
	usertype3=rstxtt("usertype3")
	usertype4=rstxtt("usertype4")
	usertype5=rstxtt("usertype5")
	usertype6=rstxtt("usertype6")
set rstxtt=nothing				
Set rstxt = conn.Execute("select usertype,jifensum,yucunkuan,Memo2 from dk501_user where userid='"&request.cookies("buyok")("userid")&"'") 
usertype=rstxt("usertype")
jifensum=rstxt("jifensum")
Memo2=rstxt("Memo2")
yucunkuan=formatnum(rstxt("yucunkuan"),2)
set rstxt=nothing
	if usertype="1" then
		response.write "<font color=red>"&usertype1&"</font>"
	elseif usertype="2" then
		response.write "<font color=red>"&usertype2&"</font>"
	elseif usertype="3" then
		response.write "<font color=red>"&usertype3&"</font>"
	elseif usertype="4" then
		response.write "<font color=red>"&usertype4&"</font>"
	elseif usertype="5" then
		response.write "<font color=red>"&usertype5&"</font>"
	elseif usertype="6" then
		response.write "<font color=red>"&usertype6&"</font>"
	end if
%>
                    <font color="red"><%=request.cookies("buyok")("userid")%></font> 我的帐户？ <span class="b-all size_14 red" style="padding-top:6px; padding-bottom:6px; padding-left:10px;"> <a href="/login/logout.asp"><font color="#0058ad"> [安全退出]</font></a></span></td>
                </tr>
                <tr>
                  <td><ul>
                    <li style="padding-top:6px; padding-bottom:6px; padding-left:10px;" class="b-all size_14 red"> 积分：<font color="#ff0000"><b><%=jifensum%></b></font> 帐户余额：<font color="#ff0000"><b><%=yucunkuan%>元</b></font></li>
                  </ul>
                  </td>
                </tr>
                <tr>
                  <td>管理员通告：<%=Memo2%></td>
                </tr>
                <tr></tr>
              </table>
	          <br></td>
  </tr>
</table>

</div>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>
