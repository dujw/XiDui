<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<%
action=request("action")
if action="" then
%>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;商品高级搜索</font></B></font></TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F" width="99%">
  <form name=search action=manage_product2.asp method=post>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>根据商品类别</td>
      <td height="25"><%	Dim a,b,c 
	set rs=server.createobject("adodb.recordset") 
	sql="select distinct LarCode,LarSeq from dk501_productclass order by LarSeq" 
	rs.Open sql,conn,1,1 
	response.write "<select name='fenlei'>"
	response.write	"<option value='' selected>选择要搜索的商品分类</option>"
	if rs.bof and rs.eof then
	response.write "<option selected value=''>-暂无分类-</option>"
	else
	Do While Not rs.eof	
	response.write "<option value='"&rs("larCode")&"'>"&rs("LarCode")&"</option>"
	rs.movenext
	  if rs.eof then exit do
	loop
	response.write "<option value='未归类'>所有未归类商品</option>"
	end if
	rs.close
	set rs=nothing
%>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center> 根据商品名称 </td>
      <td height="25"><input type=text name=name_key size=40 maxlength=20></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>根据商品简介 </td>
      <td height="25"><input type=text name=jj_key size=40 maxlength=20></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>根据商品内容 </td>
      <td height="25"><input type=text name=nr_key size=40 maxlength=20></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>根据商品价格 </td>
      <td height="25">价格范围：从
        <input type=text name=jg1_key size=5 maxlength=4>
        元 至
        <input type=text name=jg2_key size=5 maxlength=4>
        元</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>根据是否推荐 </td>
      <td height="25"><input type="radio" name="tj_key" value="1">
        是
        <input type="radio" name="tj_key" value="2">
        否
        <input type="radio" name="tj_key" value="3">
        全部</td>
    </TR>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>根据是否特价 </td>
      <td height="25"><input type="radio" name="tej_key" value="1">
        是
        <input type="radio" name="tej_key" value="2">
        否
        <input type="radio" name="tej_key" value="3">
        全部</td>
    </TR>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>根据商品状态 </td>
      <td height="25"><input type="radio" name="zt_key" value="1">
        显示
        <input type="radio" name="zt_key" value="2">
        关闭
        <input type="radio" name="zt_key" value="3">
        全部 </td>
    </TR>
    <tr bgcolor="#FFFFFF">
      <td width=20% height="25" align=center>搜索结果排序方式 </td>
      <td height="25"><input type="radio" name="paixu" value="1" checked>
        时间降
        <input type="radio" name="paixu" value="2">
        时间升
        <input type="radio" name="paixu" value="3">
        编号降
        <input type="radio" name="paixu" value="4">
        编号升
        <input type="radio" name="paixu" value="5">
        名称
        <input type="radio" name="paixu" value="6">
        点击数 </td>
    </TR>
    <tr bgcolor="#FFFFFF">
      <td colspan=2><INPUT name="action" TYPE="hidden" value="ok">
        <INPUT name=action TYPE="submit" value="开始搜索"></td>
    </tr>
  </form>
</table>
<%
else

'开始搜索
if trim(request("jg1_key"))<>"" then
  if not isNumeric(trim(request("jg1_key"))) then
	response.write "<script language='javascript'>"
	response.write "alert('您填写的内容不符合要求，价格必须是数字，请检查！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
  end if
end if

if trim(request("jg2_key"))<>"" then
  if not isNumeric(trim(request("jg2_key"))) then
	response.write "<script language='javascript'>"
	response.write "alert('您填写的内容不符合要求，价格必须是数字，请检查！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
  end if
end if

	if request("name_key")<>"" then addsql=addsql+" and prodname like '%"&request("name_key")&"%'"
	if request("jj_key")<>"" then addsql=addsql+" and ProdDisc like '%"&request("jj_key")&"%'"
	if request("nr_key")<>"" then addsql=addsql+" and MemoSpec like '%"&request("nr_key")&"%'"
	if request("tj_key")="1" then addsql=addsql+" and remark='1'"
	if request("tj_key")="2" then addsql=addsql+" and remark='0'"
	if request("zt_key")="1" then addsql=addsql+" and online=true"
	if request("zt_key")="2" then addsql=addsql+" and online=false"
	if request("tej_key")="1" then addsql=addsql+" and tejia='1'"
	if request("tej_key")="2" then addsql=addsql+" and tejia='0'"
	if request("fenlei")<>"" then addsql=addsql+" and larcode='"&request.form("fenlei")&"'"
	if trim(request("jg1_key"))<>"" then addsql=addsql+" and PriceList>"&request("jg1_key")
	if trim(request("jg2_key"))<>"" then addsql=addsql+" and PriceList<"&request("jg2_key")
	if request("paixu")="1" then orderby=" order by adddate desc"
	if request("paixu")="2" then orderby=" order by adddate"
	if request("paixu")="3" then orderby=" order by prodid desc"
	if request("paixu")="4" then orderby=" order by prodid"
	if request("paixu")="5" then orderby=" order by prodname"
	if request("paixu")="6" then orderby=" order by ClickTimes desc"

if addsql="" then
response.cookies("buyok")("sql")=""
Response.Redirect("manage_product2.asp")
else

sql="select * from bproduc where dk501_product is not null" + addsql +orderby
response.cookies("buyok")("sql")=sql
Response.Redirect("manage_prod1.asp")
end if

end if
%>
</BODY>
</html>
