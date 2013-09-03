<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
<div align="left">
<%
set rs = server.createobject("adodb.recordset")
sql = "select  top 300 RecAddress2,RecPhone2,Recname2,OrderTime from dk501_borderList where kuaidi='顺风' and Status='66' and del<>1 order by ID desc"
rs.open sql,conn,1,1
if not rs.eof then
	do while not rs.eof
		Response.Write rs("RecAddress2")&"<br>"
		Response.Write rs("Recname2")&"<br>"
		Response.Write rs("RecPhone2")&"<br><br>"
	rs.movenext
	loop
	rs.close
	set rs=Nothing
else
	response.Write "还没有您想要的你要打印的数据，请等候！"
end if
%>
</div>
</body>
</html>

