<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<%
action = request("action")
if action="" then
	call showclass()
end if
'编辑类别名字
if action="editclass" then
	call editclass()
end if
'删除二级分类
if action="delMidclass" then
	MidCode=request("MidCode")
	LarCode=request("LarCode")
	call delmidclass()
end if
'增加二级分类
if action="addMidclass" then
call addmidclass()
end if

if action="addLarclass" then
'增加一级分类
call addlarclass()
end if
'删除一级分类
if action="delLarclass" then
	LarCode=request("LarCode")
	call dellarclass()
end if


sub addmidclass()
'增加二级分类
LarCode=trim(request("LarCode"))
if request("add")="ok" then
newclass=trim(request("newclass"))
If trim(newclass)="" or instr(newclass,"&")>0 or instr(newclass,"%")>0 or instr(newclass,"'")>0 or instr(newclass,"&quot;")>0 then
response.write "<script language='javascript'>"
response.write "alert('出错了，资料填写不完整或不符合要求，请检查后重新提交。');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
response.end
end if
set rs=conn.execute("select * from dk501_productclass where MidCode='"&newclass&"' and LarCode='"&LarCode&"'")
if not (rs.eof and rs.bof) then
response.write "<script language='javascript'>"
response.write "alert('出错了，已经有一个同名二级分类存在，请使用其它名称吧。');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
set rs=nothing
set rs=nothing
sql="select * from dk501_productclass where LarCode='"&LarCode&"' order by MidSeq"
set rs=conn.execute(sql)
while Not rs.eof
counter=rs("MidSeq")+1
larseq=rs("LarSeq")
rs.movenext
wend
set rsadd=server.createobject("adodb.recordset")
rsadd.open "dk501_productclass",conn,1,3
rsadd.addnew
rsadd("LarSeq")=larseq
rsadd("LarCode")=LarCode
rsadd("MidSeq")=counter
rsadd("MidCode")=newclass
rsadd.update
rsadd.close
set rsadd=nothing
rs.close
set rs=nothing
response.write "<script language='javascript'>"
response.write "alert('操作成功，已添加了二级分类“"&trim(request("newclass"))&"”。');"
response.write "location.href='manage_product.asp';"
response.write "</script>"
response.end
else
%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">　<B><font color="#FF0000">&nbsp;商品分类管理</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
		  <table height="5"><tr><td></td></tr></table>

<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
<tr bgcolor="#FFFFFF"><td height=25 colspan=2 class=td>&nbsp;添加二级分类</td>
</tr>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center>一级分类名称</td><td><%=LarCode%></td></tr>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center valign=top">现有二级分类名称</td><td>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from dk501_productclass where LarCode='"&LarCode&"' order by MidSeq asc"
rs.open sql,conn,1,1
if not (rs.eof and rs.bof) then
do while not rs.eof
response.write rs("midcode")&"<br>"
rs.movenext
loop
end if
%>
</td></tr>
<form action='manage_product.asp?action=addMidclass' method=post name=addmidclass>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center>新增二级分类名称</td><td><INPUT TYPE='text' NAME='newclass' size=20 maxlength=20 value=''>[请勿输入<font color=red>% &quot; ' &amp; + = \</font>等禁用字符]</td></tr>
<tr bgcolor="#FFFFFF"><td colspan=2><INPUT TYPE='hidden' name=add value='ok'><INPUT TYPE='hidden' name=LarCode value='<%=LarCode%>'><INPUT TYPE='submit' value='保存设置'>
</td></tr>
</form>
</table>
<table height="5"><tr><td></td></tr></table>

<%
rs.close
set rs=nothing
end if
end sub

sub addlarclass()
'增加一级分类
if request("add")="ok" then
	If trim(request("newclass"))="" or instr(request("newclass"),"&")>0 or instr(request("newclass"),"%")>0 or instr(request("newclass"),"'")>0 or instr(request("newclass"),"&quot;")>0 then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，资料填写不完整或不符合要求，请检查后重新提交。');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if
	set rs=conn.execute("select * from dk501_productclass where LarCode='"&trim(request("newclass"))&"'")
		if not (rs.eof and rs.bof) then
			response.write "<script language='javascript'>"
			response.write "alert('出错了，已经有一个同名分类存在，请使用其它名称。');"
			response.write "location.href='javascript:history.go(-1)';"
			response.write "</script>"
			response.end
		end if
	set rs=nothing
	set rs=conn.execute("select * from dk501_productclass order by larseq desc")
	if not (rs.eof and rs.bof) then
	count=clng(rs("larseq"))+1
else
	count=1
end if
set rs=nothing
set rsadd=Server.CreateObject("ADODB.Recordset")
sql="SELECT * FROM dk501_productclass"
rsadd.open sql,conn,3,3
rsadd.addnew
rsadd("LarSeq")=count
rsadd("MidSeq")=1
rsadd("LarCode")=trim(request.form("newclass"))
rsadd("MidCode")=trim(request.form("newclass"))
rsadd.update
rsadd.close
set rsadd=nothing
response.write "<script language='javascript'>"
response.write "alert('操作成功，已添加一级分类“"&trim(request.form("newclass"))&"”，及一个同名的二级分类。');"
response.write "location.href='manage_product.asp';"
response.write "</script>"
response.end
else
set rs=server.createobject("adodb.recordset")
sql="select distinct LarCode,LarSeq from dk501_productclass order by LarSeq asc"
rs.open sql,conn,1,1
%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid">
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">　<B><font color="#FF0000">&nbsp;商品分类管理</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>

<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
<form action='manage_product.asp?action=addLarclass' method=post name=addlarclass>
<tr bgcolor="#FFFFFF"><td height=25 colspan=2 class=td>&nbsp;添加一级分类</td>
</tr>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center>一级分类名称</td><td><INPUT TYPE='text' NAME='newclass' size=20 maxlength=20 value=''>[请勿输入<font color=red>% &quot; ' &amp; + = \</font>等禁用字符]</td></tr>
<tr bgcolor="#FFFFFF"><td colspan=2><INPUT TYPE='hidden' name=add value='ok'><INPUT TYPE='submit' value='保存设置'>
</td></tr>
</form>
</table>
<%
rs.close
set rs=nothing
end if
end sub


sub delmidclass()
'删除二级分类
conn.execute "delete from dk501_productclass where MidCode='"&request("MidCode")&"' and  larCode='"&request("larCode")&"'" 
'同步删除该分类下的所有商品
if tree_del_yesorno="1" then
conn.execute "delete from dk501_product where MidCode='"&request("MidCode")&"' and  larCode='"&request("larCode")&"'" 
else
conn.execute "update dk501_product set MidCode='未归类',Larcode='未归类' where MidCode='"&request("MidCode")&"' and  larCode='"&request("larCode")&"'" 
end if

conn.close
response.write "<script language='javascript'>"
if tree_del_yesorno="1" then
response.write "alert('操作成功，所选二级分类已被删除\n\n注意：该分类下的所有商品已被同时删除。');"
else
response.write "alert('操作成功，所选二级分类已被删除\n\n注意：该分类下的所有商品已被设为未归类商品。');"
end if
response.write "location.href='manage_product.asp';"	
response.write "</script>"
end sub

sub delLarclass()
'删除一级分类
conn.execute "delete from dk501_productclass where LarCode='"&request("LarCode")&"'" 
'同步删除该分类下的所有商品
if tree_del_yesorno="1" then
conn.execute "delete from dk501_product where LarCode='"&request("LarCode")&"'" 
else
conn.execute "update dk501_product set MidCode='未归类',Larcode='未归类' where larCode='"&request("larCode")&"'" 
end if


conn.close
response.write "<script language='javascript'>"
if tree_del_yesorno="1" then
response.write "alert('操作成功，所选一级分类已被删除\n\n注意：该分类下的所有商品已被同时删除。');"
else
response.write "alert('操作成功，所选一级分类已被删除\n\n注意：该分类下的所有商品已被设为未归类商品。');"
end if
response.write "location.href='manage_product.asp';"	
response.write "</script>"
response.end
end sub

sub editclass()
'编辑商品分类
if request("MidCode")<>"" then
 sql="select * from dk501_productclass where MidCode='"&request("MidCode")&"' and  larCode='"&request("larCode")&"'"
 text="MidCode"
elseif request("LarCode")<>"" then
 sql="select * from dk501_productclass where LarCode='"&request("LarCode")&"'"
 text="LarCode"
end if

if request("edit")="ok" then
'修改二级分类名称
if request("MidCode")<>"" then

'重名处理，检查相同的一级分类下面，是否已经有这个名称的二级分类存在
set rs=conn.execute("select * from dk501_productclass where MidCode='"&request("MidCode")&"' and  larCode='"&request("larCode")&"'")
if request("old")=request("midcode") then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，新旧名称相同，未作修改。');"
	response.write "location.href='manage_product.asp';"
	response.write "</script>"
	response.end
end if

if not ( rs.eof and rs.bof ) then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，已经有这个名称的二级分类，请使用其它名称。');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
set rs=nothing
'更新类别名称
conn.execute("update dk501_productclass set MidCode='"&trim(request("MidCode"))&"' where Midcode='"&request("old")&"' and LarCode='"&request("larcode")&"'")
'同步更新商品资料中的类别
conn.execute("update dk501_product set MidCode='"&trim(request("MidCode"))&"' where Midcode='"&request("old")&"' and LarCode='"&request("larcode")&"'")


'修改一级分类名称
elseif request("LarCode")<>"" then

set rs=conn.execute("select * from dk501_productclass where larCode='"&request("larCode")&"'")
if request("old")=request("larcode") then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，新旧名称相同，未作修改。');"
	response.write "location.href='manage_product.asp';"
	response.write "</script>"
	response.end
end if

'重名处理，检查是否已经有同名的一级分类存在
if not ( rs.eof and rs.bof ) then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，已经有这个名称的一级分类，请使用其它名称。');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
set rs=nothing

'更新类别名称
conn.execute("update dk501_productclass set LarCode='"&trim(request("LarCode"))&"' where LarCode='"&request("old")&"'")
'同步更新商品资料中的类别
conn.execute("update dk501_product set LarCode='"&trim(request("LarCode"))&"' where LarCode='"&request("old")&"'")

end if
conn.close

response.write "<script language='javascript'>"
response.write "alert('操作成功，商品类别由“"&request("old")&"”更新为“"&request(text)&"”');"
response.write "location.href='manage_product.asp';"
response.write "</script>"
response.end
else
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if text="MidCode" then
%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">　<B><font color="#FF0000">&nbsp;商品分类管理</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
		  <table height="5"><tr><td></td></tr></table>

<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
<form name=editform action='manage_product.asp?action=editclass' method=post>
<tr bgcolor="#FFFFFF"><td height=25 colspan=2 class=td>&nbsp;编辑二级分类名称</td>
</tr>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center>一级分类名称</td><td> <%=request("larCode")%> </td></tr>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center>二级分类名称</td><td><INPUT TYPE='text' NAME=MidCode size=20 maxlength=20 value='<%=rs("MidCode")%>'>[请勿输入<font color=red>% &quot; ' &amp; + = \</font>等禁用字符]</td></tr>
<tr bgcolor="#FFFFFF"><td colspan=2><INPUT TYPE='hidden' name=old value='<%=rs("MidCode")%>'><INPUT TYPE='hidden' name="larcode" value='<%=rs("larcode")%>'><INPUT TYPE='hidden' name=edit value=ok><INPUT TYPE='submit' name=submit value='保存修改'>
</td></tr>
</form>
</table>
<%
end if
if text="LarCode" then
%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">　<B><font color="#FF0000">&nbsp;商品分类设置</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
		  <table height="5"><tr><td></td></tr></table>

<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
<form name=editform action='manage_product.asp?action=editclass' method=post>
<tr bgcolor="#FFFFFF"><td height=25 colspan=2 class=td>&nbsp;编辑一级分类名称</td>
</tr>
<tr bgcolor="#FFFFFF"><td width=20% height=25 align=center>一级分类名称</td><td><INPUT TYPE='text' NAME=LarCode size=20 maxlength=20 value='<%=rs("LarCode")%>'>[请勿输入<font color=red>% &quot; ' &amp; + = \</font>等禁用字符]</td></tr>
<tr bgcolor="#FFFFFF"><td colspan=2><INPUT TYPE='hidden' name=old value='<%=rs("LarCode")%>'><INPUT TYPE='hidden' name=edit value=ok><INPUT TYPE='submit' name=submit value='提交'>
</td></tr>
</form>
</table>
<%
end if
rs.close
set rs=nothing
end if
end sub


sub showclass()
'分类管理首页
%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">　<B><font color="#FF0000">&nbsp;商品分类管理</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30><font color="#000000">&nbsp;</font></TD>
    </TR>
  </TBODY>
</TABLE>
		  <table height="5"><tr><td></td></tr></table>

<table width="95%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
<%
set rs=server.createobject("adodb.recordset")
sql="select * from dk501_productclass order by LarSeq asc, MidSeq asc"
rs.open sql,conn,1,3
dim old
if rs.eof and rs.bof then 
response.Redirect "manage_product.asp?action=addLarclass"
else
Do While Not rs.eof
if rs("LarCode")<>old then    '判断另一个一级分类，old的值在二级分类列表后赋值
%>
<tr><td width=50% height=30 bgcolor="#FFFFFF">&nbsp;<img src="../Images/small_left.gif" width="8" height="8" border=0 ../Images/small_left.gif>&nbsp;<font color=red><b><%=rs("LarCode")%></b></font>
<a alt="在<font color=red> <%=rs("LarCode")%> </font>下添加二级分类" href='manage_product.asp?action=addMidclass&LarCode=<%=rs("LarCode")%>'>（添加二级分类）</a></td>
<td bgcolor="#FFFFFF"> <u alt="编辑 <font color=red><%=rs("LarCode")%></font> 的名称" style="cursor:hand" onclick="{location.href='manage_product.asp?action=editclass&LarCode=<%=rs("LarCode")%>';}">编辑</u>&nbsp;&nbsp;<a href='manage_product.asp?action=larup&larcode=<%=rs("LarCode")%>'><img src="../Images/small_up1.gif" alt='一级向上移动' width="11" height="12" border=0 ../Images/small_up1.gif></a>&nbsp;&nbsp;<a href='manage_product.asp?action=lardown&larcode=<%=rs("LarCode")%>'><img src="../Images/small_down1.gif" alt='一级分类向下移动' width="11" height="12" border=0 ../Images/small_down1.gif></a>&nbsp;&nbsp;<u 
<%if tree_del_yesorno="1" then %>
alt="删除 <font color=red><%=rs("LarCode")%></font> 分类<br>注意：此操作将同时删除该分类下的全部商品<br>提示：您可到 <font color=red>参数设置</font> 中去看看……" 
<% else %>
alt="删除 <font color=red><%=rs("LarCode")%></font> 分类<br>注意：删除分类后，所属商品将转为未归类商品<br>提示：您可到 <font color=red>参数设置</font> 去看看……" 
<%end if%>
style="cursor:hand" value="删除" onclick="{if(confirm('注意：该操作不可恢复！\n\n确实删除所选一级分类吗？')){location.href='manage_product.asp?action=delLarclass&LarCode=<%=rs("LarCode")%>';}}">删除</u>   &nbsp; <a href='manage_product.asp?action=view&larcode=<%=rs("LarCode")%>' alt="浏览<font color=red> <%=rs("LarCode")%> </font>分类下的全部商品">浏览</a>
<%if addlar<>"no" then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=manage_product.asp?action=addLarclass><font color=blue><b>添加一级分类</b></font></a>"%></td></tr>
<%
addlar="no"
end if
%>
<tr><td width=50% height=25 bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;<font color=#0066cc>
<b><%=rs("MidCode")%></b></font></td>
<td bgcolor="#FFFFFF"><u alt="修改 <font color=red><%=rs("MidCode")%></font> 的名称" style="cursor:hand" onclick="{location.href='manage_product.asp?action=editclass&MidCode=<%=rs("MidCode")%>&larcode=<%=rs("LarCode")%>';}">修改</u>&nbsp;&nbsp;<a href='manage_product.asp?action=midup&MidCode=<%=rs("MidCode")%>&larcode=<%=rs("LarCode")%>'><img src="../Images/small_up2.gif" alt='二级分类向上移动' width="11" height="12" border=0 ../Images/small_up2.gif></a>&nbsp;&nbsp;<a href='manage_product.asp?action=middown&MidCode=<%=rs("MidCode")%>&larcode=<%=rs("LarCode")%>'><img src="../Images/small_down2.gif" alt='二级分类向下移动' width="11" height="12" border=0 ../Images/small_down2.gif></a>&nbsp;&nbsp;<u 
<%if tree_del_yesorno="1" then%>
alt="删除 <font color=red><%=rs("MidCode")%></font> 分类<br>注意：此操作将同时删除该分类下的全部商品<br>提示：您可到 <font color=red>参数设置</font> 中去看看……" 
<%else%>
alt="删除 <font color=red><%=rs("MidCode")%></font> 分类<br>注意：删除分类后，所属商品将转为未归类商品<br>提示：您可到 <font color=red>参数设置</font> 去看看……" 
<%end if%>
style="cursor:hand" onclick="{if(confirm('注意：该操作不可恢复！\n\n确实删除所选二级分类吗？')){location.href='manage_product.asp?action=delMidclass&MidCode=<%=rs("MidCode")%>&larcode=<%=rs("LarCode")%>';}}">删除</u> &nbsp; <a href='manage_product.asp?action=view&MidCode=<%=rs("MidCode")%>&larcode=<%=rs("LarCode")%>' alt="浏览<font color=red> <%=rs("MidCode")%> </font>分类下的全部商品">浏览</a></td>
</tr>
<%
old=rs("LarCode")
rs.movenext
if rs.eof then exit do
Loop
end if
response.write "<tr><td colspan=2 bgcolor=#FFFFFF><input type='button' onClick=""javascript:location.href='manage_product.asp?action=output';"" value='更新分类' alt='修改分类后，请更新分类'></td></tr>"
rs.close
set rs=nothing
end sub

'二级分类向上提升
if action="midup" then
LarCode=request("LarCode")
MidCode=request("MidCode")
dim i,mid(100)
i=0
set rs=server.createobject("adodb.recordset")
sql="select * from dk501_productclass where larcode='"&larcode&"' order by MidSeq asc"
rs.open sql,conn,1,3
do while not rs.eof
i=i+1
if rs("midcode")=midcode then G=i
if rs.eof then exit do
rs.movenext
loop
rs.movefirst
i=0
do while not rs.eof
i=i+1
if i=g-1 then
rs("midseq")=i+1
elseif i>1 and i=g then
rs("midseq")=i-1
else
rs("midseq")=i
end if
rs.update
if rs.eof then exit do
rs.movenext
loop
response.Redirect "manage_product.asp"
rs.close
set rs=nothing
end if

'二级分类向下降
if action="middown" then
LarCode=request("LarCode")
MidCode=request("MidCode")
i=0
set rs=server.createobject("adodb.recordset")
sql="select * from dk501_productclass where larcode='"&larcode&"' order by MidSeq asc"
rs.open sql,conn,1,3
do while not rs.eof
i=i+1
if rs("midcode")=midcode then G=i
if rs.eof then exit do
rs.movenext
loop
rs.movefirst
i=0
do while not rs.eof
i=i+1
if i=g then
rs("midseq")=i+1
elseif i=g+1 then
rs("midseq")=i-1
else
rs("midseq")=i
end if
rs.update
if rs.eof then exit do
rs.movenext
loop
response.Redirect "manage_product.asp"
rs.close
set rs=nothing
end if

'一级分类向上提升
if action="larup" then
LarCode=request("LarCode")
i=0
set rs=server.createobject("adodb.recordset")
sql="select * from dk501_productclass order by LarSeq asc, MidSeq asc"
rs.open sql,conn,1,3
old=""
do while not rs.eof
if rs("larcode")<>old then i=i+1
if rs("larcode")=larcode then
G=i
exit do
end if
if rs.eof then exit do
old=rs("larcode")
rs.movenext
loop
rs.movefirst
i=0
old=""
do while not rs.eof
if rs("larcode")<>old then i=i+1
if i=g-1 then
rs("larseq")=i+1
elseif i>1 and i=g then
rs("larseq")=i-1
else
rs("larseq")=i
end if
rs.update
if rs.eof then exit do
old=rs("larcode")
rs.movenext
loop
response.Redirect "manage_product.asp"
rs.close
set rs=nothing
end if

'一级分类向下降
if action="lardown" then
LarCode=request("LarCode")
i=0
set rs=server.createobject("adodb.recordset")
sql="select * from dk501_productclass order by LarSeq asc, MidSeq asc"
rs.open sql,conn,1,3
old=""
do while not rs.eof
if rs("larcode")<>old then i=i+1
if rs("larcode")=larcode then
G=i
exit do
end if
if rs.eof then exit do
old=rs("larcode")
rs.movenext
loop
rs.movefirst
i=0
old=""
do while not rs.eof
if rs("larcode")<>old then i=i+1
if i=g then
rs("larseq")=i+1
elseif i=g+1 then
rs("larseq")=i-1
else
rs("larseq")=i
end if
rs.update
if rs.eof then exit do
old=rs("larcode")
rs.movenext
loop
response.Redirect "manage_product.asp"
rs.close
set rs=nothing
end if


'更新分类
if action="output" then

    Set Fso = CreateObject("Scripting.FileSystemObject")
	set dk501_productclass=Fso.CreateTextFile(Server.MapPath("../include/class.asp"),true)


	set rs=conn.execute("select * from dk501_productclass order by LarSeq asc, MidSeq asc")
	if rs.eof and rs.bof then 
		dk501_productclass.writeline "暂无商品分类"
	else
		Do While Not rs.eof
		if rs("LarCode")<>old then
			dk501_productclass.writeline "lar:"&rs("LarCode")
		end if
			dk501_productclass.writeline rs("MidCode")
			old=rs("LarCode")
		rs.movenext
		if rs.eof then exit do
		Loop
	end if
	set rs=nothing

	dk501_productclass.close
    set fso=nothing

response.write "<script language='javascript'>"
response.write "alert('操作成功，商品类别已更新，首页将显示更新后的商品分类。');"
response.write "location.href='manage_product.asp';"
response.write "</script>"
response.end
end if

if action="view" then
if request("midcode")<>"" then
sql="select * from dk501_product where ProdId is not null and larcode='"&request("larcode")&"' and midcode='"&request("midcode")&"'"
else
sql="select * from dk501_product where ProdId is not null and larcode='"&request("larcode")&"'"
end if
response.cookies("buyok")("sql")=sql
Response.Redirect("manage_prod1.asp")
end if
%>
		  <table height="5"><tr><td></td></tr></table>
