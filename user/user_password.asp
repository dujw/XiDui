<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../include/connect.asp"-->
<!--#include file="../include/md5.asp"-->
<%
call aspsql()
'�ж��û��Ƿ��½
call checklogin()
Name   = request.cookies("buyok")("userid")
if request("edit")="ok" then call edit()

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
                  <td height="25" bgcolor="E43A23">&nbsp;<span class="input_search">�����޸�</span></td>
                </tr>
                <tr>
                  <td height="5"></td>
                </tr>
                <tr>
                  <td height="5">
				  <table width="690" align="center">
  <tr>
    <td align="center" valign="middle"><TABLE cellSpacing=0 cellPadding=0 width=690 border=0 style="BORDER-RIGHT: #AFAFAF 1px solid; BORDER-TOP: #AFAFAF 1px solid; BORDER-LEFT: #AFAFAF 1px solid; BORDER-BOTTOM: #AFAFAF 1px solid">      
        <TR>
          <TD bgColor=#f4f5fd height=25><font color="#000000">��<B>��Ա�޸�����</B></font></TD>
        </TR>
      </TABLE></td>
  </tr>
  <tr>
    <td align="center" valign="middle">
	<%
set rs=conn.execute("select UserID from dk501_user where UserId='"&request.cookies("buyok")("userid")&"'")
%>
	<form name="myinfo" action="user_password.asp" method="post">
      <table width="690" border="0" cellpadding="0" cellspacing="0" style="BORDER-RIGHT: #AFAFAF 1px solid; BORDER-TOP: #AFAFAF 1px solid; BORDER-LEFT: #AFAFAF 1px solid; BORDER-BOTTOM: #AFAFAF 1px solid">
        <TR id=ps1 vAlign=center align=left>
          <TD height=35>&nbsp; ��Ա��ν��<font color="#FF0000"><b><%=rs("UserID")%></b></font> </TD>
        </TR>
        <TR id=ps1 vAlign=center align=left>
          <TD height=35><input type=hidden name="userid" value="<%=rs("UserID")%>">
                <FONT color=#ff0000>* </FONT>��Ա�ɵ�¼���룺
            <input name="oldpassword"  type="password" class=uset_input id="oldpassword3">
          </TD>
        </TR>
        <TR id=ps1 vAlign=center align=left>
          <TD height=35><FONT color=#ff0000>* </FONT>��Ա�µ�¼���룺
            <input name="pw1" type="password" class=uset_input id="pw13">
          </TD>
        </TR>
        <TR id=ps2 vAlign=center align=left>
          <TD height=35><FONT color=#ff0000>* </FONT>ȷ���µ�¼���룺
            <input name="pw2" type="password" class=uset_input id="pw23">
          </TD>
        </TR>
        <TR align=middle>
          <TD height=35><input name="Submit" type="submit" id="Submit" class=button value="ȷ���޸�">
            &nbsp;&nbsp;&nbsp;&nbsp;
            <input type=hidden name=edit value="ok"></TD>
        </TR>
      </table>
    </form></td>
  </tr>
</table>
				  </td>
                </tr>
              </table>
      </td>
  </tr>
</table>

</div>
<!--#include file="../include/login/bootom.asp"-->
</body>
</html>
<%
sub edit()
	oldpassword=trim(request("oldpassword"))
	Pw1=trim(request("pw1"))
	Pw2=trim(request("pw2"))
	
	if oldpassword="" or Pw1="" or pw2=""  then 
		response.write "<script language='javascript'>"
		response.write "alert('�����ˣ���д��������������ԭ���뼰�����롣');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if	
	
	if Pw1<>pw2 then 
		response.write "<script language='javascript'>"
		response.write "alert('�����ˣ���������������벻����');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	
	if llen(pw1)<4 then 
		response.write "<script language='javascript'>"
		response.write "alert('�����ˣ��������������ĳ��Ȳ�����Ҫ�����4λ��');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	
	set rs=conn.execute("select * from buser where UserId='"&request.cookies("buyok")("userid")&"'")
	if rs("userpassword")<>md5(oldpassword) then 
		response.write "<script language='javascript'>"
		response.write "alert('�����ˣ��������ԭ���벻��ȷ��');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if	
	
	if ucase(request.cookies("buyok")("userid"))<>ucase(request.form("userid")) then
		response.write "<script language='javascript'>"
		response.write "alert('�����ˣ�����Ȩ���д˲�����');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	
	sql = "select userpassword from dk501_user where UserId='"&request.cookies("buyok")("userid")&"'"
	set rs=Server.Createobject("ADODB.RecordSet")
	rs.Open sql,conn,1,3
	rs("userpassword")=md5(pw1)
	rs.update
	rs.close
	set rs=nothing
	
	response.write "<script language='javascript'>"
	response.write "alert('�����ɹ������Ѿ��ɹ��޸ĵ�½���롣\n\n���μ��µ����룡');"
	response.write "location.href='user_password.asp';"
	response.write "</script>"
	response.end
end sub	
conn.close
set conn=nothing
%>
