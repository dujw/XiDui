<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
<TR> 
      <TD height="30" align="left" nowrap bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000"> &nbsp;����ͳ��</font></B></font></TD>
    </TR>
</TABLE>		  
<table height="5"><tr><td></td></tr></table>

<table dk501_border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F" width="99%">
  <tr bgcolor="#FFFFFF"> 
    <td height=25 align="center">�������</td>
    <td height="25" align=center nowrap>��������</td>
    <td height="25" align=center nowrap>��Ʒ���</td>
    <td height="25" align=center>��Ʒ����</td>
    <td align=center nowrap>��ݷ�ʽ</td>
    <td align=center nowrap>��������</td>
    <td align=center nowrap>�����С</td>
    <td height="25" align=center nowrap>����</td>
    <td height="25" align=center nowrap>�ɱ�</td>
    <td height="25" align=center nowrap>���۽��</td>
  </tr>
  
<%
Recname2=Trim(request("Recname2"))
Product_Factory=Trim(request("StatusDefine"))

sql="select A.BuyPrice,A.ProdId,A.ordertime,A.Userid,A.xinhao,A.ProdUnit,B.kuaidi,B.OrderNum,C.ProdName,C.Product_Price,C.Product_number,C.Product_Factory from dk501_border A,dk501_borderList B,dk501_product C where A.OrderNum<>'' and A.OrderNum = B.OrderNum and A.ProdId=C.Prodid and b.Status='99' and b.del<>1"
	if Recname2 <> "" then 
		sql = sql & " and Success_time = '"&Recname2 &"'"
	else
		sql = sql & " and Success_time = '"&date()&"'"
	end if
	if Product_Factory <> "" then 
		sql = sql & " and Product_Factory = '"&Product_Factory&"'"
	end if
	sql= sql+"  order by a.ordertime desc"
set rs=Server.Createobject("ADODB.RecordSet")
rs.open sql,conn,1,3
if rs.eof and rs.bof  then
	response.write "<tr  bgcolor=#FFFFFF><td height=25 colspan=10 align=center>��ʱû�ж���</td></tr>"
	response.write "<form name=searchsoft method=get action=manage_order7.asp>"
  	response.write "<tr bgcolor=#FFFFFF>"
    response.write "<td colspan=10 height=30><strong>&nbsp;�������ݣ�</strong>"
     response.write " <input name=Recname2 type=text class=smallInput id=Recname2 size=20 maxlength=50 value="&Request("Recname2")&">"
      response.write "<input name=Query type=submit id=Query value='�� ѯ'>"
      response.write "����ʱ�� "&date()&"</td>"
 response.write "</tr>"
  response.write "</form>"
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
	Sum = FormatNumber(csng(rs("BuyPrice"))*csng(rs("ProdUnit")),2)
	Sum = FormatNumber(Sum,2)
	Total = Sum + Total
	%>
	<tr bgcolor=#FFFFFF>
	<td height=22 align="center"><a alt='�鿴/�༭����' href=manage_order1.asp?action=list&ID=<%=rs("OrderNum")%>><%=rs("ordernum")%></a></td>
	<td height="22" align=center nowrap><%=rs("ordertime")%></td>
	<td height="22" align="center" nowrap><%=rs("ProdId")%></td>
	<td height="22"><a alt='�鿴����Ʒ��ǰ̨ҳ��' href=../gotourl.asp?list.asp?ProdId=<%=rs("ProdId")%> target='blank_'><%=rs("ProdName")%></a></td>
	<td height="22" align="center"><%=rs("kuaidi")%></td>
	<td height="22" align="center" bgcolor="#CAE8F9"><%=rs("ProdUnit")%></td>
	<td height="22" align="center" bgcolor="#F0D0E7"><%=rs("xinhao")%></td>
	<td height="22" align=center bgcolor="#FCFAD1"><%=rs("Product_number")%></td>
	<td height="22" align=center bgcolor="#D6EBF3"><%=rs("Product_Price")*rs("ProdUnit")%></td>
	<td height="22" align=center><%=FormatNumber(csng(rs("BuyPrice"))*csng(rs("ProdUnit")),2)%></td>
	</tr>
	<%
	pages = pages - 1
	rs.movenext
if rs.eof then exit do
loop
%>
  <tr bgcolor=#FFFFFF>
	<td height=20 colspan=10 align=right>���죺
	  <%
	set rs1 = server.createobject("adodb.recordset")
	sql1 = "select sum(youfei) from dk501_borderList where  Status='99' and del <> 1"
		if Recname2 <> "" then 
			sql1 = sql1 & " and Success_time = '"&Recname2&"'"
		else
			sql1 = sql1 & " and Success_time = '"&date()&"'"
		end if	
	rs1.open sql1,conn,1,1
	if not (rs1.eof and rs1.bof) then
		Response.Write "���ʷ��ǣ�"&rs1(0)&"&nbsp;Ԫ&nbsp;"
	else
		Response.Write "���ʷ��ǣ�0&nbsp;Ԫ&nbsp;"
	end if
	%>
	��ҳ�ܼ����ۣ�<%=formatnumber(total,2)%> &nbsp; </td>
  </tr>
  <form name="searchsoft" method="get" action="manage_order7.asp">
  <tr bgcolor="#FFFFFF">
    <td colspan=10 height=30><strong>&nbsp;�������ݣ�</strong>
      <input name="Recname2" type="text" class=smallInput id="Recname2" size="20" maxlength="50" value="<%=Trim(Request("Recname2"))%>">
      <input name="Query" type="submit" id="Query" value="�� ѯ">
      ����ʱ�� <%=date()%></td>
  </tr>
  </form>
</table>

<%
call listpage()
end if
rs.close
set rs=nothing

sub listpage()
	response.write "�ܼ�¼��<font color=red>"&totalrs&"</font> &nbsp; ÿҳ<font color=red>20</font>&nbsp; "
if page = 1 then
	response.write "<font color=darkgray>��ҳ ǰҳ</font>"
else
	response.write "<a href=manage_order7.asp?Recname2="&Recname2&"&page=1>��ҳ</a> <a href=manage_order7.asp?Recname2="&Recname2&"&page="&page-1&">ǰҳ</a>"
end if

action="manage_order7.asp?StatusDefine="&Product_Factory&"&Recname2="&Recname2&""				
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
	response.write "<font color=darkgray> ��ҳ ĩҳ</font>"
else
	response.write " <a href=manage_order7.asp?StatusDefine="&Product_Factory&"&Recname2="&Recname2&"&page="&page+1&">��ҳ</a> <a href=manage_order7.asp?StatusDefine="&Product_Factory&"&Recname2="&Recname2&"&page="&allpages&">ĩҳ</a>"
end if
end sub
%>
</body>
</html>
