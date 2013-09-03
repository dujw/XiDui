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
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;常规信息设置</font></B></font></TD>
  </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="99%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
<form action=manage_System1.asp method=post name=setup>
    <tr> 
      <td width=20% align=right bgcolor="#FFFFFF">网站开关状态 &nbsp;</td>
      <td bgcolor="#FFFFFF">&nbsp;
        <input type="radio" name="kaiguan" value="1" <%if rs("kaiguan")<>0 then%>checked<%end if%>>
开放 &nbsp;
<input type="radio" name="kaiguan" value="0" <%if rs("kaiguan")=0 then%>checked<%end if%>>
关闭</td>
    </TR>
	    <tr> 
      <td width=20% align=right bgcolor="#FFFFFF">网站关闭时的提示语 &nbsp;</td>
      <td bgcolor="#FFFFFF"><input name="guanbi" type="text" style="overflow:auto;" value="<%=rs("guanbi")%>" size="60"></td>
    </TR>
    <tr> 
      <td width=20% align=right bgcolor="#FFFFFF">网站名称 &nbsp;</td>
      <td bgcolor="#FFFFFF"> <input name="sitename" type="text" value="<%=rs("sitename")%>" size="60"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="必填项。请填写网站名称<br>合理的填写名称，有利于搜索引擎找到你的网站">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>网站描述 &nbsp;</td>
      <td> <input name="sitedescription" type="text" value="<%=rs("sitedescription")%>" size="60"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="选填项。填写网站的简短描述，便于搜索引擎找到你的网站">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>网站关键字 &nbsp;</td>
      <td> <input name="sitekeywords" type="text" value="<%=rs("sitekeywords")%>" size="60"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="选填项。填写网站关键字，便于搜索引擎找到你的网站<br>多个关键词用西文逗号隔开">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>网站域名 &nbsp;</td>
      <td> <input name="siteurl" type="text" value="<%=Trim(rs("siteurl"))%>" size="60"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="必填项。填写网站域名，请以http://开头，结尾不要加/">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>电子邮箱 &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_mail")%>" size=60 name=adm_mail maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="必填项。用于接收订单通知等"> </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>公司名称 &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_comp")%>" size=60 name=adm_comp maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="选填项，可空出<br>此内容显示在主页的“联系方式”页面">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>联系地址 &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_address")%>" size=60 name=adm_address maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="选填项，可空出<br>此内容显示在主页的“联系方式”页面">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>邮政编码 &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_post")%>" size=60 name=adm_post maxlength="6"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="选填项，可空出<br>此内容显示在主页的“联系方式”页面">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>联 系 人 &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_name")%>" size=60 name=adm_name maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="选填项，可空出<br>此内容显示在主页的“联系方式”页面">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>联系电话 &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_tel")%>" size=60 name=adm_tel maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="选填项，可空出<br>此内容显示在主页的“联系方式”页面">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>传真号码 &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_fax")%>" size=60 name=adm_fax maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="选填项，可空出<br>此内容显示在主页的“联系方式”页面">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% align=right>移动电话 &nbsp;</td>
      <td> <input type=text value="<%=rs("adm_mob")%>" size=60 name=adm_mob axlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="选填项，可空出<br>此内容显示在主页的“联系方式”页面">      </td>
    </TR>
    <tr bgcolor="#FFFFFF"> 
      <td align=right>ICP备案号 &nbsp;</td>
      <td><input name=adm_icp type=text id="adm_icp" value="<%=rs("adm_icp")%>" size=60 maxlength="50"> 
        &nbsp;&nbsp; <img src=../Images/folderclosed.gif alt="选填项，可空出<br>此内容显示在主页的“联系方式”页面">      </td>
    </TR>

    <tr bgcolor="#FFFFFF"> 
      <td colspan=2><INPUT name="ok" TYPE="hidden" value="ok"> <INPUT NAME="adm_kf" TYPE="hidden" value="1"> 
        <INPUT TYPE="hidden" value="blue" name="kf_color"> <INPUT TYPE="hidden" value="<%=rs("help_hang")%>" name="help_hang"> 
        <INPUT TYPE="hidden" name="reg" value="1"> <INPUT TYPE="hidden" name="bbs" value="1"> 
        <INPUT name=action TYPE="submit" value="保存设置"></td>
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
		response.write "alert('出错了，请填写网站关闭时的提示语！');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	if request.form("sitename")="" or request.form("siteurl")="" or request.form("adm_mail")=""  then
		response.write "<script language='javascript'>"
		response.write "alert('出错了，资料填写不完整或有错误！');"
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
response.write "alert('操作成功，您设置的信息已经保存！');"
response.write "location.href='"&url&"';"
response.write "</script>"
end if%>