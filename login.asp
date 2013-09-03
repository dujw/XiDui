<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/connect.asp" -->
<%
'	Dim check
'	Randomize timer
'	check=Int((899999)*Rnd +100000)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="../image/web_css/login.css" rel="stylesheet" type="text/css" />
<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
</head>
<script language="JavaScript">
function checkform(){
	if (document.loginbox.userid.value.length ==0){
		alert("提示：请输入您的用户名！");
		document.loginbox.userid.focus();
		return false;
	}
	if (document.loginbox.password.value.length==0){
		alert("提示：请输入您的密码！");
		document.loginbox.password.focus();
		return false;
	}
//	if (document.loginbox.check1.value.length==0){
//		alert("提示：请输入您的正确验证码字！");
//		document.loginbox.check1.focus();
//		return false;
//	}
	return true;
}	
</script>
<body>
<!--#include file="include/web_top.asp" -->
<div id='div_Public'>
  <div id='mid_login'>
    <table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
      <form method="post" name="loginbox" action="ckeckuser.asp" onSubmit="return checkform();">
        <tr>
          <td height="50" colspan="2" valign="bottom" style="border-bottom-width: 2px;border-bottom-style: solid;border-bottom-color: #010101;"><img src="image/login/mid/denglu.jpg" width="102" height="25" /></td>
        </tr>
        <tr>
          <td width="188" height="40" align="right">用户名：</td>
          <td width="512" height="40"><input name="userid" type="text" id="userid" size="30" />
            请填写有效的下次登录的用户名</td>
        </tr>
        <tr>
          <td height="40" align="right">请设定密码：</td>
          <td height="40"><input name="password" type="password" id="password" size="33" />
            密码请设为6-16位字母或数字</td>
        </tr>
<!--        <tr>
          <td height="40" align="right">请输入验证码：</td>
          <td height="40"><input name="check1" type="text" size="6" maxlength="6" />
            <font color="#C04411"><%'=check%></font>
            <input type="hidden" name="check2" value="<%'=check%>" />
            请输入左边的验证码!
            </td>
        </tr>-->
        <tr>
          <td height="40" align="right">&nbsp;</td>
          <td height="40"><input type="hidden" name="login" value="ok">
            <input type="hidden" name="cook" value="0"><input type="image" src="image/login/mid/dengl.jpg" /></td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="center">有任何疑问请点击 <a href="Contact.asp?id=12">帮助中心</a> 或 <a href="Contact.asp?id=15">联系客服</a></td>
        </tr>
      </form>
    </table>
  </div>
</div>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>
