<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<%
action=request("ok")
if action="" then 
Set rs = conn.Execute("select kaiguan,guanbi,sitename,sitedescription,sitekeywords,siteurl,adm_mail,adm_comp,adm_address,adm_post,adm_name,adm_tel,adm_fax,adm_mob,adm_icp,help_hang,menu from dk501_bconfig") 
%>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;������Ϣ����</font></B></font></TD>
  </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="99%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
<form action=manage_System1.asp method=post name=setup>
    <tr> 
      <td width=20% align=right bgcolor="#FFFFFF">��վ����״̬ &nbsp;</td>
      <td bgcolor="#FFFFFF">&nbsp;
        <input type="radio" name="kaiguan" value="1" <%if rs("kaiguan")<>0 then%>checked<%end if%>>
���� &nbsp;
<input type="radio" name="kaiguan" value="0" <%if rs("kaiguan")=0 then%>checked<%end if%>>
�ر�</td>
    </TR>
	    <tr> 
      <td width=20% align=right bgcolor="#FFFFFF">��վ�ر�ʱ����ʾ�� &nbsp;</td>
      <td bgcolor="#FFFFFF"><input name="guanbi" type="text" style="overflow:auto;" value="<%=rs("guanbi")%>" size="60"></td>
    </TR>
    <tr> 
      <td width=20% align=right bgcolor="#FFFFFF">��վ���� &nbsp;</td>
      <td bgcolor="#FFFFFF"> <input name="sitename" type="text" value="<%=rs("sitename")%>" size="60"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="���������д��վ����<br>�������д���ƣ����������������ҵ������վ">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>��վ���� &nbsp;</td>
      <td> <input name="sitedescription" type="text" value="<%=rs("sitedescription")%>" size="60"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="ѡ�����д��վ�ļ���������������������ҵ������վ">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>��վ�ؼ��� &nbsp;</td>
      <td> <input name="sitekeywords" type="text" value="<%=rs("sitekeywords")%>" size="60"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="ѡ�����д��վ�ؼ��֣��������������ҵ������վ<br>����ؼ��������Ķ��Ÿ���">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>��վ���� &nbsp;</td>
      <td> <input name="siteurl" type="text" value="<%=Trim(rs("siteurl"))%>" size="60"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="�������д��վ����������http://��ͷ����β��Ҫ��/">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>�������� &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_mail")%>" size=60 name=adm_mail maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="��������ڽ��ն���֪ͨ��"> </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>��˾���� &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_comp")%>" size=60 name=adm_comp maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="ѡ����ɿճ�<br>��������ʾ����ҳ�ġ���ϵ��ʽ��ҳ��">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>��ϵ��ַ &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_address")%>" size=60 name=adm_address maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="ѡ����ɿճ�<br>��������ʾ����ҳ�ġ���ϵ��ʽ��ҳ��">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>�������� &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_post")%>" size=60 name=adm_post maxlength="6"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="ѡ����ɿճ�<br>��������ʾ����ҳ�ġ���ϵ��ʽ��ҳ��">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>�� ϵ �� &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_name")%>" size=60 name=adm_name maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="ѡ����ɿճ�<br>��������ʾ����ҳ�ġ���ϵ��ʽ��ҳ��">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>��ϵ�绰 &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_tel")%>" size=60 name=adm_tel maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="ѡ����ɿճ�<br>��������ʾ����ҳ�ġ���ϵ��ʽ��ҳ��">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>������� &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_fax")%>" size=60 name=adm_fax maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="ѡ����ɿճ�<br>��������ʾ����ҳ�ġ���ϵ��ʽ��ҳ��">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>�ƶ��绰 &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_mob")%>" size=60 name=adm_mob axlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="ѡ����ɿճ�<br>��������ʾ����ҳ�ġ���ϵ��ʽ��ҳ��">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td align=right>ICP������ &nbsp;</td>
      <td><input name=adm_icp type=text id="adm_icp" value="<%=rs("adm_icp")%>" size=60 maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="ѡ����ɿճ�<br>��������ʾ����ҳ�ġ���ϵ��ʽ��ҳ��">      </td>
    </TR>

    <tr bgcolor="#FFFFFF"> 
      <td colspan=2><INPUT name="ok" TYPE="hidden" value="ok"> <INPUT NAME="adm_kf" TYPE="hidden" value="1"> 
        <INPUT TYPE="hidden" value="blue" name="kf_color"> <INPUT TYPE="hidden" value="<%=rs("help_hang")%>" name="help_hang"> 
        <INPUT TYPE="hidden" name="reg" value="1"> <INPUT TYPE="hidden" name="bbs" value="1"> 
        <INPUT name=action TYPE="submit" value="��������"></td>
    </tr>
	</form>
</table>
</BODY>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
if action="ok" then
	if request.form("kaiguan")="0" and request.form("guanbi")="" then
		response.write "<script language='javascript'>"
		response.write "alert('�����ˣ�����д��վ�ر�ʱ����ʾ�');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	if request.form("sitename")="" or request.form("siteurl")="" or request.form("adm_mail")=""  then
		response.write "<script language='javascript'>"
		response.write "alert('�����ˣ�������д���������д���');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from dk501_bconfig"
rs.open sql,conn,1,3
	rs("guanbi")=request.form("guanbi")
	rs("kaiguan")=request.form("kaiguan")
	rs("sitename")=trim(request.form("sitename"))
	rs("sitedescription")=trim(request.form("sitedescription"))
	rs("sitekeywords")=trim(request.form("sitekeywords"))
	rs("siteurl")=trim(request.form("siteurl"))
	rs("adm_mail")=trim(request.form("adm_mail"))
	rs("adm_comp")=request.form("adm_comp")
	rs("adm_address")=request.form("adm_address")
	rs("adm_post")=request.form("adm_post")
	rs("adm_name")=request.form("adm_name")
	rs("adm_tel")=request.form("adm_tel")
	rs("adm_fax")=request.form("adm_fax")
	rs("adm_mob")=request.form("adm_mob")
	rs("adm_icp")=request.form("adm_icp")
	
	rs.update
url="manage_System1.asp"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script language='javascript'>"
response.write "alert('�����ɹ��������õ���Ϣ�Ѿ����棡');"
response.write "location.href='"&url&"';"
response.write "</script>"
end if%>