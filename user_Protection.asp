<!--#include file="include/connect.asp" -->
<%
call aspsql()
'判断用户是否登陆
call checklogin()
Name   = request.cookies("buyok")("userid")
if request("set")="ok" then call setq()
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
    <td width="260" valign="top"><!--#include file="left.asp"--></td>
    <td valign="top">
	          <table width="690" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" bgcolor="E43A23">&nbsp;<span class="input_search">申请密保</span></td>
                </tr>
                <tr>
                  <td height="5"></td>
                </tr>
                <tr>
                  <td height="5">
				  <table width="650" align=center>
  <tr>
    <td align="center" valign="middle"><TABLE cellSpacing=0 cellPadding=0 width=690 border=0 style="BORDER-RIGHT: #AFAFAF 1px solid; BORDER-TOP: #AFAFAF 1px solid; BORDER-LEFT: #AFAFAF 1px solid; BORDER-BOTTOM: #AFAFAF 1px solid">
        <TR>
          <TD bgColor=#f4f5fd height=25><font color="#000000">　<B>设置密码保护问题</B></font></TD>
        </TR>
    </TABLE></td>
  </tr>
  <tr>
    <td align="center" valign="middle">	
<%
set rs=conn.execute("select UserID,UserQuestion,UserAnswer from dk501_user where UserId='"&request.cookies("buyok")("userid")&"'")
%> 
        <form name="myinfo" action="user_Protection.asp" method="post">
          <table width="690" border="0" cellpadding="0" cellspacing="0" style="BORDER-RIGHT: #AFAFAF 1px solid; BORDER-TOP: #AFAFAF 1px solid; BORDER-LEFT: #AFAFAF 1px solid; BORDER-BOTTOM: #AFAFAF 1px solid">
            <TR id=ps1 vAlign=center align=left>
              <TD width="650" height=35>&nbsp; 会员称谓：<font color="#FF0000"><b><%=rs("UserID")%></b></font>
                  <input type=hidden name="userid" value="<%=rs("UserID")%>"></TD>
            </TR>
            <TR id=ps1 vAlign=center align=left>
              <TD width="650" height=35><FONT color=#ff0000>&nbsp;* </FONT>密码提示问题：
                <input size=35 type="text" name="UserQuestion" value="<%=rs("UserQuestion")%>" maxlength=50 class=regtxt>
              </TD>
            </TR>
            <TR id=ps1 vAlign=center align=left>
              <TD width="650" height=35><FONT color=#ff0000>&nbsp;* </FONT>提示问题答案：
              <input size=35 type="text" name="UserAnswer1" value="" maxlength=50 class=regtxt></TD>
            </TR>
            <TR id=ps1 vAlign=center align=left>
              <TD width="650" height=35><FONT color=#ff0000>&nbsp;* </FONT>确认您的答案：
              <input size=35 type="text" name="UserAnswer2" value="" maxlength=50 class=regtxt></TD>
            </TR>
            <TR align=middle>
              <TD width="650" height=35><input name="Submit" type="submit" class=button id="Submit" value="确认修改">
                &nbsp;&nbsp;&nbsp;&nbsp;
                <input type=hidden name=set value="ok">
              </TD>
            </TR>
          </table>
        </form></td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
				  </td>
                </tr>
              </table>
      </td>
  </tr>
</table>

</div>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<%
sub setq()
	UserQuestion=trim(request("UserQuestion"))
	UserAnswer1=trim(request("UserAnswer1"))
	UserAnswer2=trim(request("UserAnswer2"))
	
	if UserQuestion="" or UserAnswer1="" or UserAnswer2=""  then 
		response.write "<script language='javascript'>"
		response.write "alert('出错了，填写不完整，请填写取回密码的问题及答案。');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if	
	
	if UserAnswer1<>UserAnswer2 then 
		response.write "<script language='javascript'>"
		response.write "alert('出错了，两次输入的答案不相同。');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	
	if len(UserAnswer1)<3 then 
		response.write "<script language='javascript'>"
		response.write "alert('出错了，您输入答案的太短了，要求最低3位。');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	
	if ucase(request.cookies("buyok")("userid"))<>ucase(request("userid")) then
		response.write "<script language='javascript'>"
		response.write "alert('出错了，您无权进行此操作。');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	
	sql = "select UserID,UserQuestion,UserAnswer from dk501_user where UserId='"&request.cookies("buyok")("userid")&"'"
	set rs=Server.Createobject("ADODB.RecordSet")
	rs.Open sql,conn,1,3
	rs("UserQuestion")=UserQuestion
	rs("UserAnswer")=md5(UserAnswer1)
	rs.update
	rs.close
	set rs=nothing	
	response.write "<script language='javascript'>"
	response.write "alert('操作成功，您设置的问题和答案已保存。\n\n在您忘记密码的时候，可通过回答此问题来找回密码！');"
	response.write "location.href='user_Protection.asp';"
	response.write "</script>"
	response.end
end sub
%>
<% call CloseConn() %>
