<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->	
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/JavaScript">

function CheckAll(form) {
 for (var i=0;i<form.elements.length;i++) {
 var e = form.elements[i];
 if (e.name != 'chkall') e.checked = form.chkall.checked; 
 }
 }
</script>
</head>
<BODY>
<%
action = request("action")
if action="删除" then
ProdNum=replace(request("id")," ","")
	call prodclose()
end if

if action="关闭" then 
ProdNum=replace(request("id")," ","")
	call prodclose()
end if

if action="打开" then
ProdNum=replace(request("id")," ","")
	call prodopen()
end if
if action="tuijian" then
	id=request("id")
	call tuijian()
end if
%>
<%
sub prodclose()
if ProdNum="" or isnull(ProdNum) then
	response.write "<script language='javascript'>"
	response.write "alert('您什么也没有选择，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
else
	Set rsclose=Server.CreateObject("ADODB.Recordset")
	sqlclose="select online from dk501_product where ProdNum in ("&ProdNum&")"
	rsclose.open sqlclose,conn,1,3
		if rsclose.eof and rsclose.bof then
			response.write "<script language='javascript'>"
			response.write "alert('该商品并不存在，单击“确定”返回上一页！');"
			response.write "location.href='javascript:history.go(-1)';"
			response.write "</script>"
			response.end
		else
			Do while not rsclose.eof
			rsclose("online")=3
			rsclose.update
			rsclose.movenext
			loop
		end if 
	response.write "<script language='javascript'>"
	response.write "alert('所选商品已经被关闭！');"
	response.write "location.href='manage_product2.asp?show=yes';"
	response.write "</script>"
	
	rsclose.close
	set rsclose=nothing
end if
end sub

sub prodopen()
if ProdNum="" or isnull(ProdNum) then
	response.write "<script language='javascript'>"
	response.write "alert('您什么也没有选择，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
else
Set rsopen=Server.CreateObject("ADODB.Recordset")
sqlopen="select online from dk501_product where ProdNum in ("&ProdNum&")"
rsopen.open sqlopen,conn,1,3
if rsopen.eof and rsopen.bof then
	response.write "<script language='javascript'>"
	response.write "alert('该商品并不存在，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
else
Do while not rsopen.eof
	rsopen("online")=2
	rsopen.update
	rsopen.movenext
loop
end if 
response.write "<script language='javascript'>"
response.write "alert('所选商品已经被打开！');"
response.write "location.href='manage_product2.asp?show=yes';"
response.write "</script>"
rsopen.close
set rsopen=nothing
end if
end sub

sub tuijian()
set rstuijian=server.createobject("adodb.recordset")
sqltuijian = "select remark,TJDate from dk501_product where ProdId='"&id&"'"
rstuijian.open sqltuijian,conn,1,3
if not (rstuijian.eof and rstuijian.bof) then
	if rstuijian("remark")="1" then 
		rstuijian("remark")="0"
		tjian="所选商品已取消推荐，单击“确定”返回商品列表。"
	else
		rstuijian("remark")="1"
		rstuijian("TJDate")=now()
		tjian="所选商品已列为推荐商品，单击“确定”返回商品列表。"
	end if
rstuijian.update
rstuijian.close
set rstuijian=nothing
end if
response.write "<script language='javascript'>"
response.write "alert('"&tjian&"');"
response.write "location.href='manage_product2.asp?show=yes';"
response.write "</script>"
end sub
%>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;添加产品</font></B></font></TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table width="99%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
<%
usersql=request.cookies("buyok")("usersql")
if request("show")="yes" then
  response.cookies("buyok")("usersql")=""
end if
pages = 10
'页面数据开始查询开始
keywords=LCase(Trim(request("keywords")))
UserType=Trim(request("UserType"))
ProdName=LCase(Trim(request("ProdName")))
'页面数据开始查询结束
set rs = server.createobject("adodb.recordset")
sql="select * from dk501_product where ProdNum is not null and online=2"
if ProdName <> "" then 
	sql = sql & " and ProdName like '%" & ProdName & "%' "
end if
sql = sql + "order by addDate desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "<tr bgcolor=#FFFFFF><td colspan=7 align=center height=50>暂时没有数据</td></tr>"
	response.end 
	rs.close
	set rs=nothing
end if
totalrs=rs.RecordCount
rs.pageSize = pages
allPages = rs.pageCount
page = Cint(Request("page"))
If not isNumeric(page) then page=1
if isEmpty(page) or int(page) <= 1 then
	page = 1
elseif int(page) > allPages then
	page = allPages 
end if
rs.AbsolutePage = page
%>
<form name=Prodlist action=manage_product2.asp method=post>
  <tr bgcolor="#f4f5fd">
    <td width=2% height="25" align=center nowrap>选</td>
    <td width=3% align=center nowrap>图片</td>
    <td height="25" align=center nowrap>编号</td>
    <td height="25" align=center nowrap>名称 [点击名称预览该商品]</td>
    <td height="25" align=center nowrap>现有尺码</td>
    <td align=center nowrap>添加日期</td>
    <td colspan="2" align=center nowrap>操作</td>
    <td align=center nowrap>状态</td>
    <td height="25" align=center nowrap>点击数</td>
  </tr>
<%
Total = 0
Do While Not rs.eof and pages>0
%>
  <tr bgcolor="#FFFFFF">
    <td height="21" align="center"><input type='checkbox' value='<%=rs("ProdNum")%>' name=id></td>
    <td height="21" align="center"><img src='../../<%=rs("ImgPrev")%>' width="50" height="50"></td>
    <td height="21" align="left"><%=rs("ProdId")%></td>
    <td height="21" align="left"><a target='_blank' href='../gotourl.asp?/list.asp?ProdId=<%=rs("ProdId")%>' alt='查看&lt;font color=red&gt; <%=rs("ProdName")%> &lt;/font&gt;的前台页面'><%=rs("ProdName")%></a></td>
    <td height="21" align="center"><%=rs("prod1")%></td>
    <td height="21" align="center"><%=rs("addDate")%></td>
    <td height="21" align="center" bgcolor="#CAE8F9"><%response.write "<input type=button onClick='javascript:location.href=""manage_product3.asp?page="&currentpage&"&action=detail&id="&rs("ProdId")&""";' value='编辑'>"%></td>
	<%
if rs("remark")="0" then
	tjian="../Images/prodgif1.gif"
else
	tjian="../Images/prodgif3.gif"
end if
%>
    <td align="center" bgcolor="#F0D0E7"><%response.Write "<img border=0 src="&tjian&" onClick='javascript:location.href=""manage_product2.asp?action=tuijian&id="&rs("ProdId")&""";'>"%></td>
    <td height="21" align="center" nowrap><%
if rs("online")=1 then 
	response.write "等待审核"
elseif rs("online")=2 then 
	response.write "审核通过"
else
	response.write "下架等待处理"
end if
%></td>
    <td height="21" align="center"><%=rs("ClickTimes")%></td>
    </tr>
  <%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
%>	
  <tr bgcolor="#FFFFFF">
    <td colspan="1" align="center"><input type='checkbox' name=chkall onclick='CheckAll(this.form)'></td>
    <td colspan="9">全选 <input type="submit" name="action" value="删除" onClick="{if(confirm('商品删除后将无法恢复，您确定要删除选定的商品吗？')){this.document.Prodlist.submit();return true;}return false;}"> 
<input type="submit" name="action" value="关闭">
<input type="submit" name="action" value="打开">
<input type="button" name="action" onClick="javascript:location.href='manage_product1.asp';" value="添加商品"></td>
  </tr>
  </form>
  <form name="searchsoft" method="get" action="manage_product2.asp">
  <tr bgcolor="#FFFFFF">
    <td height="30" colspan="15" align="left"><strong>&nbsp;查找数据：</strong>
          <input name="ProdName" type="text" class=smallInput id="ProdName" size="30" maxlength="50">
          <input name="Query" type="submit" id="Query" value="查 询">
        &nbsp;&nbsp;请输入用户的ID名称。如果为空，则查找所有数据。</td>
  </tr>
  </form>
  <tr bgcolor="#FFFFFF">
    <td height="30" colspan="15" align="center">
	<%
rs.close
set rs=nothing
'if allpages<=1 then exit sub
	response.write "总订单数<font color=red>"&totalrs&"</font> &nbsp; 每页<font color=red>20</font>&nbsp; "
if page = 1 then
	response.write "<font color=darkgray>首页 前页</font>"
else
	response.write "<a href=manage_product2.asp?Recname2="&Recname2&"&show=yes&page=1>首页</a> <a href=manage_product2.asp?Recname2="&Recname2&"&show=yes&page="&page-1&">前页</a>"
end if


action="manage_product2.asp?Recname2="&Recname2&"&show=yes"				
 				 if page =< 5 then
							for i=1 to 10								
									if page=i then
										response.Write "["&i&"]"& vbCrLf
									else
										if i =< allpages then
											Response.Write("<A HREF=" & action & "&Page="&i&">"&i&"</A> " & vbCrLf)
										end if													
									end if								
							next	
					else
						Response.Write("<A HREF=" & action & "&Page=1>1...</A> " & vbCrLf)	
						cc=5
						for i=page-4 to page+cc
							if page=i then
								response.Write "["&i&"]"& vbCrLf								
							else  			
								if i=< allpages then			
									Response.Write("<A HREF=" & action & "&Page="&i&">"&i&"</A> " & vbCrLf)	
								end if 
							end if 						
						next													
					end if 

if page = allpages then
	response.write "<font color=darkgray> 下页 末页</font>"
else
	response.write " <a href=manage_product2.asp?Recname2="&Recname2&"&show=yes&page="&page+1&">下页</a> <a href=manage_product2.asp?Recname2="&Recname2&"&show=yes&page="&allpages&">末页</a>"
end if
	response.write " &nbsp;&nbsp;&nbsp;第"&page&"页 共"&allpages&"页"

%>	</td>
  </tr>
</table>
</BODY>
</html>
<% call CloseConn() %>
