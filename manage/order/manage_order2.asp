<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<!--#include file="../../include/md5.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
    <TR> 
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">　<B><font color="#FF0000">搜索订单</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>

<%
action=request("action")
if action="" then
%>
<TABLE width=99% border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
<form name=search action=manage_order2.asp method=post>
  <tr bgcolor="#FFFFFF">
    <td width=20% height="25" align=center>根据订单状态</td>
    <td height="25"><select name='Status'>
        <option value=''>请选择订单状态</option>
        <%
	set rs=conn.execute("select * from dk501_types Order by Status asc")
	do while not rs.eof
	response.write "<option value='"&rs("Status")&"'>"&rs("StatusDefines")&"</option>"
	rs.movenext
	if rs.eof then exit do
	loop
	set rs=nothing
%>
      </select></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td width=20% height="25" align=center> 根据订单号 </td>
    <td height="25"><input type=text name=ordernum size=25 maxlength=12></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td width=20% height="25" align=center> 根据提交订单的会员ID </td>
    <td height="25"><input type=text name=userid size=25 maxlength=16></td>
  </tr>
    <tr bgcolor="#FFFFFF">

  <td width=20% height="25" align=center>根据收货人姓名 </td>
  <td height="25"><input type=text name=recname size=25 maxlength=20></td></tr>
    <tr bgcolor="#FFFFFF">

  <td width=20% height="25" align=center>根据收货人电话 </td>
  <td height="25"><input type=text name=RecPhone size=25 maxlength=20></td></tr>
   <!-- <tr bgcolor="#FFFFFF">

  <td width=20% height="25" align=center>根据移动电话 </td>
  <td height="25"><input type=text name=CompPhone size=25 maxlength=20></td></tr>-->
    <tr bgcolor="#FFFFFF">

  <td width=20% height="25" align=center>根据联系地址 </td>
  <td height="25"><input type=text name=RecAddress size=25 maxlength=100></td></tr>
    <tr bgcolor="#FFFFFF">

  <td width=20% height="25" align=center>根据邮政编码 </td>
  <td height="25"><input type=text name=ZipCode size=25 maxlength=8></td></tr>
    <tr bgcolor="#FFFFFF">

  <td width=20% height="25" align=center>根据提交订单时间 </td>
  <td height="25"><input type="radio" name="ordertime" value="1">
    今天 &nbsp; <input type="radio" name="ordertime" value="2">
    昨天 &nbsp; <input type="radio" name="ordertime" value="7">
    本周 &nbsp; <input type="radio" name="ordertime" value="30">
    本月 &nbsp; <input type="radio" name="ordertime" value="all">
    全部 &nbsp; </td></tr>
  <tr bgcolor="#FFFFFF">
    <td height="25" colspan=2><INPUT name="action" TYPE="hidden" value="ok">
      <INPUT name=action TYPE="submit" value="开始搜索"></td>
  </tr>
  </form>
</table>
<br><font color=red>※提示：</font>
支持模糊搜索，可根据一个或多个条件进行搜索，非搜索条件可空出。 
<%
else
if trim(request("ordernum"))<>"" then
  if not isNumeric(trim(request("ordernum"))) then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，订单号必须是数字！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
  end if
end if
	ordersql=""
	if request("Status")<>"" then ordersql=ordersql+" and b.Status='"&request("Status")&"'"
	if trim(request("ordernum"))<>"" then ordersql=ordersql+" and a.ordernum like '%"&trim(request("ordernum"))&"%'"
	if trim(request("userID"))<>"" then ordersql=ordersql+" and a.userid like '%"&trim(request("userid"))&"%'"
	if trim(request("recname"))<>"" then ordersql=ordersql+" and a.recname2 like '%"&trim(request("recname"))&"%'"	
	if trim(request("RecAddress"))<>"" then ordersql=ordersql+" and a.RecAddress2 like '%"&trim(request("RecAddress"))&"%'"
	if trim(request("RecPhone"))<>"" then ordersql=ordersql+" and a.RecPhone2 like '%"&trim(request("RecPhone"))&"%'"
	if trim(request("CompPhone"))<>"" then ordersql=ordersql+" and a.CompPhone like '%"&trim(request("CompPhone"))&"%'"
	if trim(request("ZipCode"))<>"" then ordersql=ordersql+" and a.ZipCode2 like '%"&trim(request("ZipCode"))&"%'"

	'if request("ordertime")="1" then  ordersql=ordersql+" and a.ordertime >= #"&date()&"#"
	if request("ordertime")="1" then  ordersql=ordersql+" and datediff(day,[ordertime],getdate())=0"
	if request("ordertime")="2" then  ordersql=ordersql+" and datediff(day,[ordertime],getdate())=1"
	if request("ordertime")="7" then  ordersql=ordersql+" and datediff(week,[ordertime],getdate())=0"
	if request("ordertime")="30" then  ordersql=ordersql+" and datediff(month,[ordertime],getdate())=0"
	if request("ordertime")="all" then  ordersql=ordersql
	
response.cookies("buyok")("ordersql")=ordersql
Response.Write ordersql&"+"&request.cookies("buyok")("ordersql")
if ordersql="" then
Response.Redirect("manage_order2.asp")
else
Response.Redirect("manage_order1.asp")
end if

end if
%>
</body>
</html>
