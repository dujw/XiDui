<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../include/connect.asp"-->
<%
call aspsql()
'判断用户是否登陆
call checklogin()
Name   = request.cookies("buyok")("userid")

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="../image/web_css/login.css" rel="stylesheet" type="text/css" />
</head>
<body>
<!--#include file="../include/top.asp"-->
<div id="div_Public">
<table width="950" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="260" valign="top"><!--#include file="left.asp"--></td>
    <td valign="top">
	          <table width="690" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" bgcolor="E43A23">&nbsp;<span class="input_search">用户中心</span></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td><table width="680" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" >
                    <%
set rs=conn.execute("select * from dk501_user where UserId='"&request.cookies("buyok")("userid")&"'")
%>
                    <tr valign="center" align="left">
                      <td width="52%" height="25" bgcolor="#FFFFFF">&nbsp;&nbsp; 会员称谓：<font color="#FF0000"><b><%=rs("UserID")%></b></font> </td>
                      <td width="48%" height="25" bgcolor="#FFFFFF">&nbsp;&nbsp; 会员的权限：
                        <%
									usertype=rs("usertype")									
									if usertype="2" then
										usertype=type2
									elseif usertype="3" then
										usertype=type3
									elseif usertype="4" then
										usertype=type4
									elseif usertype="5" then
										usertype=type5
									elseif usertype="6" then
										usertype=type6
									else
										usertype=type1
									end if
									response.write "<font color=#FF0000><b>"&Usertype&"</b></font>&nbsp;&nbsp;&nbsp;&nbsp;"
								%></td>
                    </tr>
                    <tr valign="center" align="left">
                      <td height="30" bgcolor="#FFFFFF"><font color="#ff0000">&nbsp;&nbsp; * </font>您的姓名：
                        <input name="Username" type="text" class="regtxt" id="Username" value="<%=rs("UserName")%>" maxlength="16" />                      </td>
                      <td height="30" bgcolor="#FFFFFF"></td>
                    </tr>
                    <tr valign="center" align="left">
                      <td height="30" bgcolor="#FFFFFF"><font color="#ff0000">&nbsp;&nbsp; *</font> 省　　份：
                        <input name="Province"  type="text" value="<%=rs("Province")%>" size="10"  maxlength="20" class="regtxt" />                      </td>
                      <td height="30" bgcolor="#FFFFFF"><font color="#ff0000">&nbsp;&nbsp; *</font> 城　　市：
                        <input name="City"  type="text" value="<%=rs("City")%>" size="10"  maxlength="20" class="regtxt" />                      </td>
                    </tr>
                    <tr valign="center" alilgn="left">
                      <td height="30" colspan="2" bgcolor="#FFFFFF"><font color="#ff0000">&nbsp;&nbsp; *</font> 地　　址：
                        <input name="Address"  type="text" value="<%=rs("Address")%>" size="80" maxlength="70" />                      </td>
                    </tr>

                    <tr valign="center" align="left">
                      <td height="30" bgcolor="#FFFFFF"><font color="#ff0000">&nbsp;&nbsp; *</font> 电　　话：
                        <input  type="text" name="HomePhone" value="<%=rs("HomePhone")%>"  maxlength="20" class="regtxt" />                      </td>
                      <td height="30" bgcolor="#FFFFFF">&nbsp;&nbsp;<font color="#ff0000">&nbsp;&nbsp; *</font> E-mail　：
                      <input  type="text" name="UserMail" value="<%=rs("UserMail")%>" maxlength="100" class="regtxt" /></td>
                    </tr>

                    <tr valign="center" align="left">
                      <td height="30" bgcolor="#FFFFFF">&nbsp;&nbsp; 腾讯QQ号：
                        <input  type="text" name="UserQQ" value="<%=rs("UserQQ")%>"  maxlength="11" class="regtxt" />                      </td>
                      <td height="30" bgcolor="#FFFFFF">&nbsp;&nbsp; 淘宝号 &nbsp;&nbsp;：
                        <input name="taobao_name" 
                  type="text" class="regtxt" id="taobao_name" value="<%=rs("taobao_name")%>"  maxlength="100" />                      </td>
                    </tr>
                    <tr align="middle">
                      <td height="30" colspan="2" bgcolor="#FFFFFF">&nbsp;</td>
                    </tr>
                    <%
rs.close
set rs=nothing
%>
                  </table></td>
                </tr>
              </table>
	          <br></td>
  </tr>
</table>

</div>
<!--#include file="../include/login/bootom.asp"-->
</body>
</html>
<% call CloseConn() %>
