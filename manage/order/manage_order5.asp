<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<!--#include file="../../include/md5.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; "><TR> 
      <TD height="30" align="left" nowrap bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;销售统计</font></B></font></TD>
    </TR>
</TABLE>
<table height="5"><tr><td></td></tr></table>

<table border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F" width="99%">
  <tr bgcolor="#FFFFFF"> 
    <td height=25 align="center">订单编号</td>
    <td height="25" align=center nowrap>销售日期</td>
    <td height="25" align=center nowrap>商品编号</td>
    <td height="25" align=center>商品名称</td>
    <td align=center>销售数量</td>
    <td align=center>尺码大小</td>
    <td height="25" align=center nowrap>货号</td>
    <td height="25" align=center nowrap>商家</td>
  </tr>
  
  <%
if Trim(request("Status"))="" then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，参数不对请刷新左边。');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
Recname2=Trim(request("Recname2"))
Product_Factory=Trim(request("StatusDefine"))
sql="select top 1000  A.ProdId,A.ordertime,A.Userid,A.xinhao,A.ProdUnit,B.OrderNum,C.ProdName,C.Product_number,C.Product_Factory from dk501_border A,dk501_borderList B,dk501_product C where A.OrderNum<>'' and A.OrderNum = B.OrderNum and A.ProdId=C.Prodid and b.Status='"&Trim(request("Status"))&"' and b.del<>1"
	if Recname2 <> "" then 
		'sql = sql & " and (A.ProdId like '%" & Recname2 & "%' or A.ordernum like '%" & Recname2 & "%')"
		sql = sql & " and (A.ProdId like '%" & Recname2 & "%')"
	end if
	if Product_Factory <> "" then 
		sql = sql & " and Product_Factory = '"&Product_Factory&"'"
	end if
sql= sql+" order by a.ordertime desc"
set rs=Server.Createobject("ADODB.RecordSet")
rs.open sql,conn,1,3
if rs.eof and rs.bof  then
	response.write "<tr  bgcolor=#FFFFFF><td height=25 colspan=9 align=center>暂时没有订单</td></tr>"
else
pages=20
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
	Total = 0
	Do While Not rs.eof and pages>0
	%>
	<tr bgcolor=#FFFFFF>
	<td height=22 align="center"><a href=manage_order1.asp?action=list&ID=<%=rs("OrderNum")%>><%=rs("ordernum")%></a></td>
	<td height="22" align=center nowrap><%=rs("ordertime")%></td>
	<td height="22" align="center" nowrap><%=rs("ProdId")%></td>
	<td height="22"><a alt='查看此商品的前台页面' href=../gotourl.asp?list.asp?ProdId=<%=rs("ProdId")%> target='blank_'><%=rs("ProdName")%></a></td>
	<td height="22" align="center" bgcolor="#CAE8F9"><%=rs("ProdUnit")%></td>
	<td height="22" align="center" bgcolor="#F0D0E7"><%=rs("xinhao")%></td>
	<td height="22" align=center bgcolor="#FCFAD1"><%=rs("Product_number")%></td>
	<td height="22" align=center bgcolor="#D6EBF3"><%=rs("Product_Factory")%></td>
	</tr>
	<%
	pages = pages - 1
	rs.movenext
if rs.eof then exit do
loop
%> 
  <form name="searchsoft" method="get" action="manage_order5.asp">
  <tr bgcolor="#FFFFFF">
    <td colspan=8 height=30><strong>&nbsp;查找数据：</strong>
      <input name="Recname2" type="text" class=smallInput id="Recname2" size="20" maxlength="50">
	  <input name="Status" type="hidden" size="20" maxlength="50" value="<%=request("Status")%>">
      <input name="Query" type="submit" id="Query" value="查 询"></td>
  </tr>
  </form>
</table>
<%
call listpage()
end if
rs.close
set rs=nothing

sub listpage()
	response.write "总记录数<font color=red>"&totalrs&"</font> &nbsp; 每页<font color=red>20</font>&nbsp; "
if page = 1 then
	response.write "<font color=darkgray>首页 前页</font>"
else
	response.write "<a href=manage_order5.asp?Status="&Trim(request("Status"))&"&StatusDefine="&Product_Factory&"&Recname2="&Recname2&"&page=1>首页</a> <a href=manage_order5.asp?StatusDefine="&Product_Factory&"&Recname2="&Recname2&"&page="&page-1&">前页</a>"
end if

action="manage_order5.asp?Status="&Trim(request("Status"))&"&StatusDefine="&Product_Factory&"&Recname2="&Recname2&""				
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
	response.write " <a href=manage_order5.asp?Status="&Trim(request("Status"))&"&StatusDefine="&Product_Factory&"&Recname2="&Recname2&"&page="&page+1&">下页</a> <a href=manage_order5.asp?StatusDefine="&Product_Factory&"&Recname2="&Recname2&"&page="&allpages&">末页</a>"
end if
end sub
%>
</body>
</html>
