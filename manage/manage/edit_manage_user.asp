<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<!--#include file="../../include/md5.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<%
if request("deladm")<>"" then
	deladm=request("deladm")
	call deladmin()
end if

admin=request("admin")
if admin="" then admin=request.cookies("buyok")("admin")
if request("edit")<>"ok" then
%>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;管理员权限设置</font></B></font></TD>
  </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
  <tr bgcolor="#FFFFFF"><td width="20%">
    <table border=0 width=90%>
      <tr>
        <td>现有管理员：</td>
      </tr>
      <%
		Set rs = conn.Execute("select * from manage_user")
		do while not rs.eof
		response.write "<tr><td><img border=0 src=../images/small_left.gif> "
		if admin=rs("username") then
		response.write "<a href=edit_manage_user.asp?admin="&rs("username")&"><font color=red><b>"&rs("username")&"</b></font></a>"
		else
		response.write "<a href=edit_manage_user.asp?admin="&rs("username")&">"&rs("username")&"</a>"
	end if
%></td>
      <td><img border=0 src=../images/delete.gif alt="删除此管理员" style="cursor:hand" onClick="{if(confirm('该操作不可恢复！\n\n确实要删除这个管理员吗？ ')){location.href='edit_manage_user.asp?deladm=<%=rs("username")%>';}}"></td>
      </tr>
      <%
	rs.movenext
	loop
	%>
    </table><td width="0%"></td>
    <td width="98%" align=center><table width=100% border=0 cellpadding="0" cellspacing="1">
      <tr><td colspan="2"><form action="edit_manage_user.asp" method=post name=manage>
    
          <tr>
            <td colspan="2"><font color=red><b><%=admin%></b></font> 的管理权限：</td>
          </tr>
          <%
	Set rs = conn.Execute("select * from manage_user where username='"&admin&"'")
		  if rs.eof and rs.bof then
			response.write "<script language='javascript'>"
			response.write "alert('没有此管理员！');"
			response.write "location.href='javascript:history.go(-1)';"
			response.write "</script>"
			response.end
		  end if
 		  manage=rs("manage")
		%>
          <tr>
            <td colspan="2"><input type=checkbox value='yes' name=manage01 <%if instr(manage,"01")>0 then%>checked<%end if%>>综合设置</td>
          </tr>
          <tr>
            <td width="4%" align="center" bgcolor="#E1E1E1">+</td>
            <td width="96%" height="25" bgcolor="#E1E1E1">
			<input type=checkbox value='yes' name=manage013 <%if instr(manage,"013")>0 then%>checked<%end if%>>常规信息设置
			<input type=checkbox value='yes' name=manage014 <%if instr(manage,"014")>0 then%>checked<%end if%>>网站帮助信息
			<input type=checkbox value='yes' name=manage018 <%if instr(manage,"018")>0 then%>checked<%end if%>>操作可以编辑			</td>
          </tr>
          <tr>
            <td colspan="2"><input type=checkbox value='yes' name=manage02 <%if instr(manage,"02")>0 then%>checked<%end if%>>广告管理</td>
          </tr>
          <tr>
            <td colspan="2"><input type=checkbox value='yes' name=manage03 <%if instr(manage,"03")>0 then%>checked<%end if%>>商品管理</td>
          </tr>
		  <tr>
            <td width="4%" align="center" bgcolor="#E1E1E1">+</td>
            <td width="96%" height="25" bgcolor="#E1E1E1">
			<input type=checkbox value='yes' name=manage031 <%if instr(manage,"031")>0 then%>checked<%end if%>>分类管理/添加商品
			<input type=checkbox value='yes' name=manage032 <%if instr(manage,"032")>0 then%>checked<%end if%>>商品管理/商品搜索 
			<input type=checkbox value='yes' name=manage033 <%if instr(manage,"033")>0 then%>checked<%end if%>>商品转移/商品图片水印
			<input type=checkbox value='yes' name=manage038 <%if instr(manage,"038")>0 then%>checked<%end if%>>操作可以编辑	 
			<input type=checkbox value='yes' name=manage039 <%if instr(manage,"039")>0 then%>checked<%end if%>>产品价格查看修改
			</td>
          </tr>
          <tr>
            <td colspan="2"><input type=checkbox value='yes' name=manage04 <%if instr(manage,"04")>0 then%>checked<%end if%>>订单管理</td>
          </tr>
		  <tr>
            <td width="4%" align="center" bgcolor="#E1E1E1">+</td>
            <td width="96%" height="25" bgcolor="#E1E1E1">
			<input type=checkbox value='yes' name=manage041 <%if instr(manage,"041")>0 then%>checked<%end if%>>订单管理/搜索/统计
			<input type=checkbox value='yes' name=manage042 <%if instr(manage,"042")>0 then%>checked<%end if%>>快递打印
			<input type=checkbox value='yes' name=manage043 <%if instr(manage,"043")>0 then%>checked<%end if%>>回收站
			<input type=checkbox value='yes' name=manage045 <%if instr(manage,"045")>0 then%>checked<%end if%>>本天销售统计
			<input type=checkbox value='yes' name=manage046 <%if instr(manage,"046")>0 then%>checked<%end if%>>销售统计|利润
			<input type=checkbox value='yes' name=manage047 <%if instr(manage,"047")>0 then%>checked<%end if%>>退换货
			<input type=checkbox value='yes' name=manage048 <%if instr(manage,"048")>0 then%>checked<%end if%>>操作可以编辑			</td>
          </tr>
          <tr>
            <td colspan="2"><input type=checkbox value='yes' name=manage05 <%if instr(manage,"05")>0 then%>checked<%end if%>>会员管理</td>
          </tr>
          <tr>
            <td colspan="2"><input type=checkbox value='yes' name=manage06 <%if instr(manage,"06")>0 then%>checked<%end if%>>新闻管理</td>
          </tr>
          <tr>
            <td colspan="2"><input type=checkbox value='yes' name=manage07 <%if instr(manage,"07")>0 then%>checked<%end if%>>投诉意见</td>
          </tr>
          <tr>
            <td colspan="2"><input type=checkbox value='yes' name=manage08 <%if instr(manage,"08")>0 then%>checked<%end if%>>保留</td>
          </tr>
          <tr>
            <td colspan="2"><input type=checkbox value='yes' name=manage09 <%if instr(manage,"09")>0 then%>checked<%end if%>>友情链接</td>
          </tr>
          <tr>
            <td colspan="2"><input type=checkbox value='yes' name=manage10 <%if instr(manage,"10")>0 then%>checked<%end if%>>安全设置</td>
          </tr>
          <tr>
            <td colspan="2"><input type=checkbox value='yes' name=manage11 <%if instr(manage,"11")>0 then%>checked<%end if%>>访问统计</td>
          </tr>
          <tr>
            <td colspan="2"><input type=checkbox value='yes' name=manage12 <%if instr(manage,"12")>0 then%>checked<%end if%>>其它信息</td>
          </tr>
    </table>
    </td>
  </tr>
  <tr bgcolor="#ffffff">
    <td colspan=3> <input type="submit" name="Submit" value="保存设置"> <input type="hidden" name="edit" value="ok"> ( 此管理员最后登录时间为：<b><font color="#FF0000"><%=rs("LoginTime")%></font></b> )
      <input type="hidden" name="admin" value="<%=admin%>"> </td>
  </tr>
</table>		
</form>
<%   
else
	manage=""
	if request.form("manage01")="yes" then manage=manage+"|01"
		if request.form("manage013")="yes" then manage=manage+"|013"
		if request.form("manage014")="yes" then manage=manage+"|014"
		if request.form("manage018")="yes" then manage=manage+"|018"
	if request.form("manage02")="yes" then manage=manage+"|02"
	if request.form("manage03")="yes" then manage=manage+"|03"
		if request.form("manage031")="yes" then manage=manage+"|031"
		if request.form("manage032")="yes" then manage=manage+"|032"
		if request.form("manage033")="yes" then manage=manage+"|033"
		if request.form("manage038")="yes" then manage=manage+"|038"
		if request.form("manage039")="yes" then manage=manage+"|039"
	if request.form("manage04")="yes" then manage=manage+"|04"
		if request.form("manage041")="yes" then manage=manage+"|041"
		if request.form("manage042")="yes" then manage=manage+"|042"
		if request.form("manage043")="yes" then manage=manage+"|043"
		if request.form("manage045")="yes" then manage=manage+"|045"
		if request.form("manage046")="yes" then manage=manage+"|046"
		if request.form("manage047")="yes" then manage=manage+"|047"
		if request.form("manage048")="yes" then manage=manage+"|048"
	if request.form("manage05")="yes" then manage=manage+"|05"
	if request.form("manage06")="yes" then manage=manage+"|06"
	if request.form("manage07")="yes" then manage=manage+"|07"
	if request.form("manage08")="yes" then manage=manage+"|08"
	if request.form("manage09")="yes" then manage=manage+"|09"
	if request.form("manage10")="yes" then manage=manage+"|10"
	if request.form("manage11")="yes" then manage=manage+"|11"
	if request.form("manage12")="yes" then manage=manage+"|12"
	
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from manage_user where username='"&admin&"'"
	rs.open sql,conn,1,3
		rs("manage")=manage
		rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	response.write "<script language='javascript'>"
	response.write "alert('管理权限设置成功！');"
	response.write "location.href='edit_manage_user.asp?admin="&admin&"';"
	response.write "</script>"
end if


sub deladmin()

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from manage_user where username='"&deladm&"'"
	rs.open sql,conn,1,3
		if (rs.bof and rs.eof) then
			response.write "<script language='javascript'>"
			response.write "alert('操作失败，没有此管理员！');"
			response.write "location.href='edit_manage_user.asp?admin="&admin&"';"
			response.write "</script>"
			response.end
		elseif deladm=request.cookies("buyok")("admin") then
			response.write "<script language='javascript'>"
			response.write "alert('您不能删除自己！');"
			response.write "location.href='edit_manage_user.asp?admin="&admin&"';"
			response.write "</script>"
			response.end
		else
		rs.delete
		rs.update
			response.write "<script language='javascript'>"
			response.write "alert('您已成功删除管理员"&deladm&"');"
			response.write "location.href='edit_manage_user.asp?admin="&admin&"';"
			response.write "</script>"
			response.end
		end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

end sub
%>
</BODY>
</html>
