<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<% call aspsql() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<%
action=request("ok")
if action="" then 
    Set rs = conn.Execute("select usertype1,usertype2,usertype3,usertype4,usertype5,usertype6 from dk501_bconfig") 
%>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;��Ա�������ۿ�����</font></B></font></TD>
  </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
  <form action=manage_user_jibie.asp method=post name=setup>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=right>��Ա����һ &nbsp; </td>
      <td height="25"> <input type=text value="<%=rs("usertype1")%>" name=usertype1  size="20" maxlength=20> 
      &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="��д��Ա��������ƣ��硰��ͨ��Ա��<br>����Ӧ����дһ����Ա����" width="18" height="18">      </td>
    </TR>
    
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=right>��Ա����� &nbsp; </td>
      <td height="25"><input name=usertype2 type=text id="usertype2" value="<%=rs("usertype2")%>" size="20" maxlength=20>        
      &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="���ޣ��ɿճ�" width="18" height="18"> </td>
    </TR>
    
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=right>��Ա������ &nbsp; </td>
      <td height="25"> <input type=text value="<%=rs("usertype3")%>" name=usertype3 size="20" maxlength=20>      </td>
    </tr>
    
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=right>��Ա������ &nbsp; </td>
      <td height="25"> <input type=text value="<%=rs("usertype4")%>" name=usertype4 size="20" maxlength=20>      </td>
    </tr>
    
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=right>��Ա������ &nbsp; </td>
      <td height="25"> <input type=text value="<%=rs("usertype5")%>" name=usertype5 size="20" maxlength=20>      </td>
    </tr>
    
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=right>��Ա������ &nbsp; </td>
      <td height="25"> <input type=text value="<%=rs("usertype6")%>" name=usertype6 size="20" maxlength=20>      </td>
    </tr>
    
    <tr bgcolor="#FFFFFF">
      <td height="25" colspan=2><INPUT name="ok" TYPE="hidden" value="ok">
        <INPUT name=action TYPE="submit" value="��������"> &nbsp;</td>
    </tr>
  </form>
</table>
	
<%
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if

if action="ok" then

	if request.form("usertype1")="" then
		response.write "<script language='javascript'>"
		response.write "alert('�����ˣ���һ����Ա����Ϊ�����');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if

	on error resume next	'�����м���������������

     Set rs=Server.CreateObject("ADODB.Recordset")
	 sql="select usertype1,usertype2,usertype3,usertype4,usertype5,usertype6 from dk501_bconfig"
	 rs.open sql,conn,1,3

 	 rs("usertype1")=trim(request.form("usertype1"))
 	 rs("usertype2")=trim(request.form("usertype2"))
 	 rs("usertype3")=trim(request.form("usertype3"))
 	 rs("usertype4")=trim(request.form("usertype4"))
 	 rs("usertype5")=trim(request.form("usertype5"))
 	 rs("usertype6")=trim(request.form("usertype6"))
	rs.update
	rs.close	
	set rs=nothing
	conn.close
	set conn=nothing

	response.write "<script language='javascript'>"
	response.write "alert('�����ɹ��������õ���Ϣ�ѱ��棡');"
	response.write "location.href='manage_user_jibie.asp';"			
	response.write "</script>"
end if%>
