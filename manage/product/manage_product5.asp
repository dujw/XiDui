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
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;商品类别转移</font></B></font></TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F" width="99%">
  <form name=search action=manage_product5.asp method=post>
    <tr bgcolor="#FFFFFF">
      <td colspan=2 class=td height=25>&nbsp;商品类别转移</td>
    </tr>
    <%	Dim a,b,c 
	set rs=server.createobject("adodb.recordset") 
	sql="select distinct LarCode,LarSeq,MidCode,midSeq from dk501_productclass order by LarSeq,MIdSeq" 
	rs.Open sql,conn,1,1 
	if rs.bof and rs.eof then
	response.write "<TR bgcolor=#FFFFFF><TD align=center colspan=2>没有对商品进行分类，不能转移商品。</td>"
	else
	response.write "<TR bgcolor=#FFFFFF><TD align=center width=100 >从</td><td>"
	response.write "<select name='fenlei1'>"
	response.write	"<option value='' selected>选择要转移的商品分类</option>"
	Do While Not rs.eof
	a=rs("larcode")	
	response.write "<option value='"&rs("larCode")&"^"&rs("MidCode")&"'>"&rs("LarCode")&" >> "&rs("MidCode")&"</option>"
	rs.movenext
	if rs.eof then exit do
	if rs("larcode")<>a then response.write "<option value=''>--------------------------</option>"
	loop
	response.write "<option value=''>--------------------------</option>"
	response.write "<option value='未归类^未归类'>所有未归类的商品</option>"

	response.write "</td><tr bgcolor=#FFFFFF><TD align=center width=100 >转移到</td><td>"

	rs.movefirst	'指针重定位
	response.write "<select name='fenlei2'>"
	response.write	"<option value='' selected>选择商品转移的目标位置</option>"
	Do While Not rs.eof
	a=rs("larcode")	
	response.write "<option value='"&rs("larCode")&"^"&rs("MidCode")&"'>"&rs("LarCode")&" >> "&rs("MidCode")&"</option>"
	rs.movenext
	if rs.eof then exit do
	if rs("larcode")<>a then response.write "<option value=''>--------------------------</option>"
	loop
	end if
	rs.close
	set rs=nothing
	response.write "</td></tr>"

%>
    <tr bgcolor="#FFFFFF">
      <td height="25" colspan=2><INPUT name=action TYPE="submit" value="开始转移">
        <INPUT name="action" TYPE="hidden" value="ok">
      </td>
    </tr>
  </form>
</table>
<%
else

'开始搜索
if request("fenlei1")="" or  request("fenlei2")="" then
	response.write "<script language='javascript'>"
	response.write "alert('您没有选择，或者没有可供选择的商品分类。');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if

if request("fenlei1")=request("fenlei2") then
	response.write "<script language='javascript'>"
	response.write "alert('选择了两个相同的分类，商品未转移。');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if


code1=split(request.form("fenlei1"),"^")
larcode1=code1(0)
MidCode1=code1(1)

code2=split(request.form("fenlei2"),"^")
larcode2=code2(0)
MidCode2=code2(1)

set rs=server.createobject("adodb.recordset")
sql = "select * from dk501_product where larcode='"&larcode1&"' and midcode='"&midcode1&"'"
rs.open sql,conn,1,3
if  (rs.eof and rs.bof) then

	response.write "<script language='javascript'>"
	response.write "alert('这个分类中没有商品可供转移。');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end

else
i=0
do while not rs.eof
rs("larcode")=larcode2
rs("midcode")=midcode2
i=i+1
if rs.eof then exit do
rs.movenext
loop

	response.write "<script language='javascript'>"
	response.write "alert('商品转移成功！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end

end if
rs.close
set rs=nothing


end if
%>
</BODY>
</html>
<% call CloseConn() %>
