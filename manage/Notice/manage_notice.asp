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
      <TD width="84%" height="30" align="left" nowrap bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;��վ֪ͨ</font></B></font></TD>
      <TD width="16%" align="right" nowrap bgColor=#f4f5fd> <a href="manage_notice_Add.asp">�����﷢��֪ͨ</a>&nbsp;</TD>
      </TR>
</TABLE>
  <form name="searchsoft" method="POST" action="manage_order8.asp">        
 <fieldset><legend>&nbsp;֪ͨ�б� &nbsp;</legend>
   <table width="99%" border="0" cellpadding="0" cellspacing="1" dk501_border="0" bgcolor="#8B979F" >
  <tr bgcolor="#FFFFFF"> 
    <td height=25 align="center">���</td>
    <td height="25" align=center nowrap>����</td>
    <td align=center nowrap>������</td>
    <td align=center nowrap>�������޸ģ�ʱ��</td>
    <td height="25" align=center nowrap>����ʱ��</td>
    <td height="25" align=center nowrap>����</td>
  </tr>
  
  <%

		SSQL="select id,Ntitle,NAddUser,NEditTime,NOverTime from DU_Notice order by id desc"
'response.write SSQL
set rs=Server.Createobject("ADODB.RecordSet")
rs.open SSQL,conn,1,3
if rs.eof and rs.bof  then
	response.write "<tr  bgcolor=#FFFFFF><td height=25 colspan=10 align=center>��ʱû�����ݣ�</td></tr>"
else
pages=20
pagesnumber=pages
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
'	'Sum = FormatNumber(csng(rs("BuyPrice"))*csng(rs("ProdUnit")),2)-FormatNumber(csng(rs("Product_Price"))*csng(rs("ProdUnit")),2)
'	Sum = FormatNumber(Sum,2)
'	Total = Sum + Total
	%>
	<tr bgcolor="#FFFFFF">
	<td height=22 align="center"><%=rs(0)%></td>
	<td height="22" align="center"><%=rs("Ntitle")%></td>
	<td height="22" align="center"><%=rs("NAddUser")%></td>
	<td height="22" align=center><%=rs("NEditTime")%></td>
	<td height="22" align=center><%=rs("NOverTime")%></td>
	<td align=center ><a href="manage_notice_Modify.asp?id=<%=rs(0)%>">�༭</a></td>
	</tr>
	<%
	pages = pages - 1
	rs.movenext
if rs.eof then exit do
loop
%>

</table>
</BR>
<%
call listpage()
end if
rs.close
set rs=nothing

FUNCTION listpage()
	response.write "¼�� <font color=red>"&totalrs&" �� </font> &nbsp; ÿҳ <font color=red> "&pagesnumber&" </font>&nbsp; "
if page = 1 then
	response.write "<font color=darkgray>��ҳ ǰҳ</font>"
else
	response.write "<a href=manage_notice.asp?page=1>��ҳ</a> <a href=manage_notice.asp?page="&page-1&">ǰҳ</a>"
end if

action="manage_notice.asp"				
 				 if page =< 5 then
							for i=1 to 10								
									if page=i then
										response.Write "["&i&"]"& vbCrLf
									else
										if i =< allpages then
											Response.Write("<A HREF=" & action & "?Page="&i&">"&i&"</A> " & vbCrLf)
										end if													
									end if								
							next	
					else
						Response.Write("<A HREF=" & action & "?Page=1>1...</A> " & vbCrLf)	
						cc=5
						for i=page-4 to page+cc
							if page=i then
								response.Write "["&i&"]"& vbCrLf								
							else  			
								if i=< allpages then			
									Response.Write("<A HREF=" & action & "?Page="&i&">"&i&"</A> " & vbCrLf)	
								end if 
							end if 						
						next													
					end if 
					
if page = allpages then
	response.write "<font color=darkgray> ��ҳ ĩҳ</font>"
else
	response.write " <a href=manage_notice.asp?page="&page+1&">��ҳ</a> <a href=manage_notice.asp?page="&allpages&">ĩҳ</a>"
end if
end FUNCTION
%> 
   </fieldset>
  </form>  

</body>
</html>
