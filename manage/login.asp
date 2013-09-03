<!--#include file="../include/connect.asp"-->
<!--#include file="../include/md5.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>网站后台登录</title>
<STYLE type="text/css">
<!--
a:link       {text-decoration: none; font-family: AdobeSm; color: #000000 }
a:visited    {text-decoration: none; color: #000000 }
A:hover      {COLOR: green; FONT-FAMILY: "宋体,MingLiU"; TEXT-DECORATION: underline}
body         {margin-top: 0; margin-bottom: 0;margin-left:0;margin-right:0;color: #1E5494;background-color: #660000}
body,td,th,div,span,input,select {font-family:宋体, fantasy;font-size:12px;}
#img_checkcode {color:#FF9900;}
#isok_checkcode {margin-left:5px;}
.logininput{border:1px solid #FFF;background:#660000;height:18px;line-height:16px;vertical-align:middle;font-family: 宋体,MingLiU, Arial;COLOR:#FFF}
.telter {FONT-SIZE: 9pt; COLOR: #dadeed; TEXT-DECORATION: none}

-->
</STYLE>
<%
	Dim check
	Randomize timer
	check=Int((899999)*Rnd +100000)
%>
</head>
<body>
<%if request("login")<>"check" then%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="600" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" height="600">
        <tr>
          <td width="33%" height="65"  align="right" valign="bottom" style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#D6E6F5', endColorStr='#6785DB', gradientType='0')"><table border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td><img src="images/back.gif"border="0"></td>
                <td><a  href='javascript:window.close()'><img src="images/close1.gif" border="0"></a></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td height="1"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="1" style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#698FD8', endColorStr='#FFFFFF', gradientType='1')"></td>
                <td style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#FFFFFF', endColorStr='#698FD8', gradientType='1')"></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td valign="middle" bgcolor="#660000"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" height="100%">
              <tr>
                <td width="49%" height="400" align="right" valign="middle"><table border="0" align="right" cellpadding="0" cellspacing="0">
                    <tr>
                      <td><img src="images/logo.png" ></td>
                      <td width="20">&nbsp;</td>
                    </tr>
                  </table></td>
                <td width="1" height="300" valign="middle"><table width="1" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="1" height="150" style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#3F63AD', endColorStr='#ffffff', gradientType='0')"></td>
                    </tr>
                    <tr>
                      <td height="150" valign="middle" style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#ffffff', endColorStr='#3F63AD', gradientType='0')"></td>
                    </tr>
                  </table></td>
                <td height="400" valign="middle"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
				<form method="post" action="login.asp" name="adminlogin">
                    <tr>
                      <td width="4%">&nbsp;</td>
                      <td width="47%" height="100%" valign="middle"><table width="254" border="0" cellpadding="2" cellspacing="2">
                          <tr>
                            <td  rowspan="2" align="center" valign="middle"><span class="telter"><strong>用户名称：</strong></span></td>
                          </tr>
                          <tr>
                            <td><INPUT  name="username"  type="text"  class="logininput" id="username"  size="22" maxLength=20></td>
                          </tr>
                          <tr>
                            <td rowspan="2" align="center" valign="middle"><span class="telter"><strong>用户密码：</strong></span></td>
                          </tr>
                          <tr>
                            <td><label>
                              <INPUT name="password" type="password"  class="logininput" id="password" size="22" maxLength=20>
                              </label></td>
                          </tr>
                          <tr>
                            <td rowspan="2" align="center" valign="middle"><span class="telter"><strong>认 证 码：</strong></span></td>
                          </tr>
                          <tr>
                            <td><label>
                              <input name="path" type="password" class="logininput" id="path" size="22" maxLength=20 />
                              </label></td>
                          </tr>
                          <tr>
                            <td width="32%" rowspan="2" align="center" valign="middle"><span class="telter"><strong>验 证 码：</strong></span></td>
                          </tr>
                          <tr>
                            <td><input name="check1" type="text" class=logininput size="6" maxlength="6" />
                              <font color="#FFFFFF"><%=check%></font>
                              <input type="hidden" name="check2" value="<%=check%>" />
							  <input type="hidden" name=login value="check">
                            </td>
                          </tr>
						  
                        </table></td>
                      <td width="49%" align="left" valign="middle"><input type="image" name="imageField" src="images/login.png" style=" width:45px; height:45px;"/></td>
                    </tr>
					</form>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td height="1"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="1" style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#698FD8', endColorStr='#FFFFFF', gradientType='1')"></td>
                <td style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#FFFFFF', endColorStr='#698FD8', gradientType='1')"></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td width="33%"  align="center" height="65" valign="bottom" style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#6785DB', endColorStr='#D6E6F5', gradientType='0')"></td>
        </tr>
      </table></td>
  </tr>
</table>
</BODY>
</HTML>
 <%
else

if session("buyok_admin_login")>=5 then
	Set rs=Server.CreateObject("ADODB.RecordSet") 
	sql="select * from dk501_bconfig"
	rs.open sql,conn,1,3
	userip=Request.serverVariables("REMOTE_ADDR")
	if instr(rs("ip"),userip)<0 then rs("ip")=rs("ip")&"@"&userip
	rs.update
	rs.close
	set rs=nothing
	response.write "<script language='javascript'>"
	response.write "alert('您涉嫌非法登陆裁难系统后台，已被系统锁定。请与技术人员联系。');"
	response.write "location.href='index.asp';"
	response.write "</script>"
	response.end
end if

check1=trim(request("check1"))   '输入的验证码
check2=trim(request("check2")) '系统自带的验证码
path=trim(request("path"))
username=trim(request("username"))
password=trim(request("password"))
adminIP=Request.ServerVariables("Remote_Addr")
	if path = "" or username="" or password="" then
		response.write "<script language='javascript'>"
		response.write "alert('填写不完整，请检查后重新提交！');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	if check1="" or check1 <> check2 then	
		Response.Write("<script language=""JavaScript"">alert(""错误：请输入正确的验证码，请返回检查！！"");history.go(-1);</script>")
	end if 

	if path <> "xiduipsd" then
		session("buyok_admin_login")=session("buyok_admin_login")+1
		response.write "<script language='javascript'>"
		response.write "alert('您填写的目录不存在，请检查后重新提交。\n\n提示：出错"&session("buyok_admin_login")&"次');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if

	set rs=conn.execute("select * from manage_user where password='"&md5(password)&"' and username='"&username&"'")
 	if not(rs.bof and rs.eof) then
		session("buyok_admin_login")=0
		Response.cookies("buyok")("admin")=username	'设置cookies
		Response.Cookies("MyCookie").Expires = DateAdd("n", 10, Now())
		conn.execute("update manage_user SET LoginTime ='"&Now()&"' WHERE username ='"&username&"'")
		Response.Redirect ("admin_index.asp")		'登入真实后台
	else	
		    response.write "<script language='javascript'>"
		    session("buyok_admin_login")=session("buyok_admin_login")+1
		    response.write "alert('您填写的用户名或者密码有误，请检查后重新输入。\n\n提示：出错"&session("buyok_admin_login")&"次');"
		    response.write "location.href='javascript:history.go(-1)';"
		    response.write "</script>"
		    response.end
	end if
	set rs=nothing
	conn.close
	set conn=nothing
end if
%>