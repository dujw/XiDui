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
		alert("��ʾ�������������û�����");
		document.loginbox.userid.focus();
		return false;
	}
	if (document.loginbox.password.value.length==0){
		alert("��ʾ���������������룡");
		document.loginbox.password.focus();
		return false;
	}
//	if (document.loginbox.check1.value.length==0){
//		alert("��ʾ��������������ȷ��֤���֣�");
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
          <td width="188" height="40" align="right">�û�����</td>
          <td width="512" height="40"><input name="userid" type="text" id="userid" size="30" />
            ����д��Ч���´ε�¼���û���</td>
        </tr>
        <tr>
          <td height="40" align="right">���趨���룺</td>
          <td height="40"><input name="password" type="password" id="password" size="33" />
            ��������Ϊ6-16λ��ĸ������</td>
        </tr>
<!--        <tr>
          <td height="40" align="right">��������֤�룺</td>
          <td height="40"><input name="check1" type="text" size="6" maxlength="6" />
            <font color="#C04411"><%'=check%></font>
            <input type="hidden" name="check2" value="<%'=check%>" />
            ��������ߵ���֤��!
            </td>
        </tr>-->
        <tr>
          <td height="40" align="right">&nbsp;</td>
          <td height="40"><input type="hidden" name="login" value="ok">
            <input type="hidden" name="cook" value="0"><input type="image" src="image/login/mid/dengl.jpg" /></td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="center">���κ��������� <a href="Contact.asp?id=12">��������</a> �� <a href="Contact.asp?id=15">��ϵ�ͷ�</a></td>
        </tr>
      </form>
    </table>
  </div>
</div>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>
