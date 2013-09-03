<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
</head>
<BODY>
<%
action=request("action")
if action="" then 
%>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER: #8B979F 1px solid; ">
    <TR>
      <TD height="30" bgColor=#f4f5fd><font color="#000000">　<B><font color="#FF0000">&nbsp;友情链接设置</font></B></font></TD>
    </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
  <form name=linklist action="manage_links_list.asp" method=post>
    <tr bgcolor="#f4f5fd" > 
      <td width="5%" valign="middle"><strong>&nbsp;选</strong></td>
      <td width="10%" valign="middle"><strong>&nbsp;序号</strong></td>
      <td width="40%" height="30" valign="middle" bgcolor="#f4f5fd"><strong>&nbsp;网站名称</strong></td>
      <td width="30%" height="30" valign="middle"><strong>&nbsp;logo</strong></td>
      <td width="15%" height="30" valign="middle"><strong>&nbsp;状态</strong></td>
    </tr>
    <%
msg_per_page = 20     '每页显示记录条数

set rs=Server.CreateObject("ADODB.recordset")
sql="select * from dk501_links order by txt desc,num asc,addtime asc"

rs.cursorlocation = 3    '设置游标
rs.pagesize = msg_per_page
rs.open sql,conn,1,1 


if rs.eof and rs.bof then
	response.write "<tr align='center' style='vertical-align: top'><td colspan='5'>暂时没有友情链接</td></tr>"
else '检测记录集是否为空
	totalrec = RS.RecordCount 'totalrec：总记录条数
	if rs.recordcount mod msg_per_page = 0 then
	n = rs.recordcount\msg_per_page 'n:总页数
	else 
	n = rs.recordcount\msg_per_page+1 
	end if 

	currentpage = request("page")
	If currentpage <> "" then
		currentpage = clng(currentpage)
		if currentpage < 1 then currentpage = 1
	else
		currentpage = 1
	End if 
	if currentpage*msg_per_page > totalrec and not((currentpage-1)*msg_per_page < totalrec) then 
		currentPage=1
	end if
	rs.absolutepage = currentpage 'absolutepage：设置指针指向某页开头
	rowcount = rs.pagesize   '设置每一页的数据记录数

	dim i
	dim k

	Do while not rs.eof and rowcount>0
	%>
    <tr bgcolor="#FFFFFF" > 
      <td height="25" class=b> <input type='checkbox' value='<%=Cstr(rs("id"))%>' name=linksid> 
      </td>
      <td height="25"><%=rs("num")%></td>
      <td height="25"><a href="manage_links_list.asp?action=modify_links&linksid=<%=rs("id")%>" alt='编辑此条友情链接'><%=rs("site")%></a> 
      </td>
      <td height="25"> <a href="../gotourl.asp?<%=rs("url")%>" target=_blank alt="去这个网站逛逛"> 
        <%
if rs("txt")="1" then
response.write "文字链接"
else
if left(rs("logo"),7)="http://" then
response.write "<img src='"&rs("logo")&"' border=0 width=88 height=31></a>"
else
response.write "<img src='../"&rs("logo")&"' border=0 width=88 height=31></a>"
end if
end if
%>
      </td>
      <td height="25"> 
        <% if rs("online")="1" then response.write "显示"%>
        <% if rs("online")="0" then response.write "<font color=red>隐藏</font>"%>
      </td>
    </tr>
    <%
	rowcount=rowcount-1
	rs.movenext
	if rs.eof then exit do
	loop	
end if     
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
    <tr bgcolor="#FFFFFF"> 
      <td height="30" colspan=7 bgcolor="#f4f5fd" class=b> <input type='checkbox' name=chkall onclick='CheckAll(this.form)'>
        全选 
        <input type="submit" name="action" value="关闭" onClick="{if(confirm('确定要在首页隐藏选定的链接吗？')){this.document.linklist.submit();return true;}return false;}"> 
        <input type="submit" name="action" value="打开" onClick="{if(confirm('确定要在首页显示选定的链接吗？')){this.document.linklist.submit();return true;}return false;}"> 
        <input type="submit" name="action" value="删除" onClick="{if(confirm('该操作不可恢复！\n\n确定删除选定的项目吗？')){this.document.linklist.submit();return true;}return false;}"> 
      <input type="button" name="action2" onClick="javascript:location.href='manage_links_list.asp?action=add_links';" value="添加友情链接"></td>
    </tr>
  </form>
</table>
		  <table height="5"><tr><td></td></tr></table>

<%call listPages()%>
<%end if

if action="add_links" then
	call add_links()
end if

if action="删除" then
delid=replace(request("id")," ","")
call delall_links()
end if

if action="关闭" then
delid=replace(request("id")," ","")
call close_links()
end if

if action="打开" then
delid=replace(request("id")," ","")
call open_links()
end if

if action="modify_links" then
	id=replace(request("id")," ","")
	call modify_links()
end if%></table>
</td></tr> </table> 
</body>
</html>

<%
sub listPages() 
if n <= 1 then exit sub 
%>
<p>>> 
<%if currentpage = 1 then%>
<font color=darkgray>首页</font>
<%else%> 
<font color=black face="arial"><a href="<%=request.ServerVariables("script_name")%>?page=1">
首页</font></a>
 <a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage-1%>">
<font color=black face="arial" >前页</a></font>
<%end if%>
<%if currentpage = n then%> 
<font color=darkgray face="arial" >后页</font>
<%else%> 
<font color=black face="arial" ><a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage+1%>">下页</a>
 <a href="<%=request.ServerVariables("script_name")%>?page=<%=n%>">末页</a></font>
<%end if%>
<font color=black face="arial" >
  页次：<%=currentpage%>/<%=n%>页&nbsp;&nbsp;每页数：<%=msg_per_page%>&nbsp;&nbsp;总数：<%=totalrec%></font></p>
<%end sub

sub delall_links()
	if request("linksid")="" or isnull(request("linksid")) then
	response.write "<script language='javascript'>"
	response.write "alert('操作失败，您什么也没有选择。');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
	else
		conn.execute("delete from dk501_links where Id in ("&request("linksid")&")")
		if err.number<>0 then '错误处理
			response.write "数据库操作失败：" & err.description
			err.clear
		else		
		 response.write "<script language=javascript>alert('所选友情链接已被删除。');location.href='manage_links_list.asp';</script>"
		end if
	end if
end sub

sub close_links()       '关闭
   if request("linksid")="" or isnull(request("linksid")) then
	response.write "<script language='javascript'>"
	response.write "alert('操作失败，您什么也没有选择。');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
   else
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from dk501_links where Id in ("&request("linksid")&")"
	rs.open sql,conn,1,3
	Do while not rs.eof
	rs("online")="0"
	rs.update
	rs.movenext
	loop			

	response.write "<script language='javascript'>"
	response.write "alert('所选友情链接已被关闭。');"
	response.write "location.href='manage_links_list.asp';"
	response.write "</script>"
   end if
end sub


sub open_links()	'打开
   if request("linksid")="" or isnull(request("linksid")) then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您什么也没有选择。');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
   else
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from dk501_links where Id in ("&request("linksid")&")"
	rs.open sql,conn,1,3
	Do while not rs.eof
	rs("online")="1"
	rs.update
	rs.movenext
	loop			

	response.write "<script language='javascript'>"
	response.write "alert('成功，所选友情链接已被打开。');"
	response.write "location.href='manage_links_list.asp';"
	response.write "</script>"
   end if
end sub


sub add_links()
if  request("send")="ok" then

'资料填写不完整
if request("site")="" or request("num")="" or request("url")="" then
	response.write "<script language='javascript'>"
	response.write "alert('友情链接的资料填写不完整，请检查后填写完整');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end if

if request.form("txt")="0" and request.form("logo")="" then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您选择了图标链接，但未填写有效的图标地址！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"	
	response.end
end if

	sql="select * from dk501_links"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3

	rs.addnew
		rs("site")=trim(request.form("site"))
	  	rs("logo")=trim(request.form("logo"))
	  	rs("url")=trim(request.form("url"))
	  	rs("num")=trim(request.form("num"))
	  	rs("online")=request.form("online")
	  	rs("txt")=request.form("txt")

	rs.update	
	rs.close
	set rs=nothing
		 response.write "<script language=javascript>alert('恭喜，已经添加了一条新的友情链接！');location.href='manage_links_list.asp';</script>"

else
%>
</p>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;添加友情连接</font></B></font></TD>
  </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
<form name="add" method="post" action="manage_links_list.asp?action=add_links">
    <tr bgcolor="#FFFFFF">
      <td  align=right>序号：</td>
      <td >
        <input name="num" type="text"  size="5" maxlength="2"> 根据序号的相对大小排列</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td  align=right>名称：</td>
      <td >
        <input name="site" type="text"  size="40" maxlength="40"> 填写网站名称</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align=right>图标：</td>
      <td>
        <input name="logo" type="text"  size="40" maxlength="100" value="http://"> 填写logo图标的网址</td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td   align=right>网址：</td>
      <td >
        <input name="url" type="text"  size="40" maxlength="100" value="http://"> 填写友情网站的网址</td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td align=right>类型：</td>
      <td>
	<input name="txt" type="radio" value='0' >图标链接
        <input name="txt" type="radio" value='1' checked>文字链接  </td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td   align=right>状态：</td>
      <td >
        <input name="online" type="radio" value="1" checked>显示  
	<input name="online" type="radio" value="0">隐藏</td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td width="10%"  align=right>&nbsp; </td>
      <td width="90%"><input type="submit" name="SubmitName" value=" 添 加 ">
          <input name="send" type="hidden"  value="ok"></td>
    </tr>
	</form>
  </table>

<%
end if
end sub

sub modify_links()
if  request.form("send")<>"ok" then
	  set rs=Server.CreateObject("adodb.recordset")
	  sql="select * from dk501_links where id="&request("linksid")
	  rs.open sql,conn,1,3

%>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;修改连接</font></B></font></TD>
  </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>
<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
<form name="modify" method="post" action="manage_links_list.asp?action=modify_links">
    <tr  bgcolor="#FFFFFF">
      <td  align=right>序号：</td>
      <td >
        <input name="num" type="text"  size="5" maxlength="2" value=<%=rs("num")%>> 根据序号的相对大小排列</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td  align=right>名称：</td>
      <td >
        <input name="site" type="text"  size="40" maxlength="40" value=<%=rs("site")%>> 填写网站名称</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td   align=right>图标：</td>
      <td >
        <input name="logo" type="text"  size="40" maxlength="100" value=<%=rs("logo")%>> 填写logo图标的网址</td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td   align=right>网址：</td>
      <td >
        <input name="url" type="text"  size="40" maxlength="100" value=<%=rs("url")%>> 填写友情网站的网址</td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td align=right>类型：</td>
      <td>
	<input name="txt" type="radio" value='0'  <% if rs("txt")="0" then%>checked<%end if%>>图标链接
        <input name="txt" type="radio" value='1' <% if rs("txt")="1" then%>checked<%end if%>>文字链接 </td> 
    </tr>

    <tr bgcolor="#FFFFFF">
      <td   align=right>状态：</td>
      <td >
        <input name="online" type="radio" value='1' <% if rs("online")="1" then%>checked<%end if%>>显示  
	<input name="online" type="radio" value='0'  <% if rs("online")="0" then%>checked<%end if%>>隐藏</td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td width="10%"  align=right>&nbsp; </td>
      <td width="90%"><input type="submit" name="SubmitName" value=" 修  改 ">
          <input name="send" type="hidden"  value="ok">
	<input name="linksid" type="hidden" value="<%=request("linksid")%>"></td>
    </tr>
	</form>
  </table>
		
<%
	  rs.close
	  set rs=nothing
else

'资料填写不完整
if request.form("site")="" or request.form("num")="" or request.form("url")="" then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，友情链接的资料填写不完整！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"	
	response.end
end if

if request.form("txt")="0" and request.form("logo")="" then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您选择了图标链接，但未填写有效的图标地址！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"	
	response.end
end if


	  set rs=Server.CreateObject("adodb.recordset")
	  sql="select * from dk501_links where id="&request("linksid")
	  rs.open sql,conn,1,3
	  if rs.eof and rs.bof then
		 response.write "<script language=javascript>alert('编号错误！');location.href='manage_links_list.asp';</script>"
		 response.end
	  else
	  	rs("logo")=trim(request.form("logo"))
	  	rs("site")=trim(request.form("site"))
	  	rs("url")=trim(request.form("url"))
	  	rs("num")=trim(request.form("num"))
	  	rs("online")=trim(request.form("online"))
	  	rs("txt")=trim(request.form("txt"))
		rs.update	
	  end if
	  rs.close
	  set rs=nothing
	 call CloseConn()
		 response.write "<script language=javascript>alert('恭喜，友情链接已经修改成功！');location.href='manage_links_list.asp';</script>"
end if
end sub%>
