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
      <TD height="30" align="left" nowrap bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;����ͳ�Ʊ���</font></B></font></TD>
      </TR>
</TABLE>
  <form name="searchsoft" method="POST" action="manage_order8.asp">        
 <fieldset><legend>&nbsp;����ѡ�� &nbsp;</legend>
   <table align="center" width="95%" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
      <tr>
	  <td width="38%" align="right" height="30"><input name="rptype" type="radio" id="rptype" value="0" <%IF Trim(Request("rptype"))="0" OR Trim(Request("rptype"))="" THEN%>checked<%END IF%>>
	    ���·�ʽ�鿴�ձ���</td>
      <td width="62%"><input name="type_d" type="text" class=smallInput id="type_d" size="6" maxlength="6" value="<%if Trim(Request("type_d"))="" then response.Write DatePart("yyyy",now())&right("0"&Month(now()),2)  else response.Write Trim(Request("type_d")) end if%>">
        �������·ݣ����磺<%=DatePart("yyyy",now())&right("0"&Month(now()),2)%> ��Ϊ�����ѯȫ�������ݣ�</td>
	  </tr>
            <tr>
	  <td width="38%" align="right" height="30"><input name="rptype" type="radio" id="rptype" value="1" <%IF Trim(Request("rptype"))="1" THEN%>checked<%END IF%>>���귽ʽ�鿴�±���</td>
      <td><input name="type_m" type="text" class=smallInput id="type_m" size="4" maxlength="4" value="<%if Trim(Request("type_m"))="" then response.Write DatePart("yyyy",now()) else response.Write Trim(Request("type_m")) end if%>">
        ��������ݣ����磺<%=DatePart("yyyy",now())%> ��Ϊ�����ѯȫ�������ݣ�</td>
	  </tr>
            <tr>
	  <td width="38%" align="right" height="30"><input type="radio" name="rptype" id="rptype" value="2" <%IF Trim(Request("rptype"))="2" THEN%>checked<%END IF%>>���귽ʽ�鿴�걨��</td>
      <td><input name="type_y" type="text" class=smallInput id="type_y" size="4" maxlength="4" value="<%if Trim(Request("type_y"))="" then response.Write DatePart("yyyy",now()) else response.Write Trim(Request("type_y")) end if%>">
        ��������ݣ����磺2012 ��Ϊ�����ѯȫ�������ݣ�</td>
	  </tr>
       <tr>
	  <td height="30" colspan="2" align="center"><input name="Query" type="submit" id="Query" value="�� �� �� ѯ"></td>
      </tr>
   </table> 
   </fieldset>
  </form>  
<table dk501_border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F" width="99%">
  <tr bgcolor="#FFFFFF"> 
    <td height=25 align="center">��������/�·�</td>
    <td height="25" align=center nowrap>���۽��(��λ��Ԫ)</td>
    <td align=center nowrap>��������(��λ����)</td>
    <td align=center nowrap>����(��λ��Ԫ)</td>
    <td height="25" align=center nowrap>�ɱ�(��λ��Ԫ)</td>
    <td height="25" align=center nowrap>�ʷ�(��λ��Ԫ)</td>
  </tr>
  
  <%
Dim rptype,type_d,type_m,type_y,SearchType,SSQL
rptype=Trim(request("rptype"))
type_d=Trim(request("type_d"))
type_m=Trim(request("type_m"))
type_y=Trim(request("type_y"))

if type_d="" then
	tiaojian_d=""
 else
 	tiaojian_d=" WHERE Buyyearmonth='"&type_d&"'"
end if

if type_m="" then
	tiaojian_m=""
 else
 	tiaojian_m=" WHERE Buyyear='"&type_m&"' "
end if

if type_y="" then
	tiaojian_y=""
 else
 	tiaojian_y=" WHERE Buyyear='"&type_y&"' "
end if

SearchType=rptype
Select Case SearchType 
	Case "0" 
		'SSQL="SELECT Success_time, sum(BuyPrice*ProdUnit) AS ZONGJINE,sum((BuyPrice-Product_Price)*ProdUnit) AS LIRUN,sum(Product_Price*ProdUnit) AS CHENGBEN ,sum(youfei) AS YOUFEI,sum(ProdUnit) AS SHULIANG  FROM RP_SALES_ALL "&tiaojian_d&" GROUP BY Success_time order by Success_time"
		 SSQL="SELECT BuyyearmonthDate,sum(ZONGJINE) AS ZONGJINE, SUM(LIRUN) AS LIRUN,SUM(CHENGBEN) AS CHENGBEN ,SUM(YOUFEI) AS YOUFEI,SUM(SHULIANG) AS SHULIANG,BuyyearmonthDate FROM RP_SALES_ALL_OrderNum "&tiaojian_d&" GROUP BY BuyyearmonthDate order by BuyyearmonthDate"
	Case "1" 
		'SSQL="SELECT Buyyearmonth, sum(BuyPrice*ProdUnit) AS ZONGJINE,sum((BuyPrice-Product_Price)*ProdUnit) AS LIRUN,sum(Product_Price*ProdUnit) AS CHENGBEN ,sum(youfei) AS YOUFEI,sum(ProdUnit) AS SHULIANG  FROM RP_SALES_ALL  "&tiaojian_m&" GROUP BY Buyyearmonth order by Buyyearmonth"
		SSQL="SELECT Buyyearmonth,sum(ZONGJINE) AS ZONGJINE, SUM(LIRUN) AS LIRUN,SUM(CHENGBEN) AS CHENGBEN ,SUM(YOUFEI) AS YOUFEI,SUM(SHULIANG) AS SHULIANG FROM RP_SALES_ALL_OrderNum "&tiaojian_m&" GROUP BY Buyyearmonth order by Buyyearmonth"
	Case "2" 
		'SSQL="SELECT Buyyear,sum(BuyPrice*ProdUnit) AS ZONGJINE,sum((BuyPrice-Product_Price)*ProdUnit) AS LIRUN,sum(Product_Price*ProdUnit) AS CHENGBEN ,sum(youfei) AS YOUFEI,sum(ProdUnit) AS SHULIANG FROM RP_SALES_ALL  "&tiaojian_y&" GROUP BY Buyyear order by Buyyear"	
		SSQL="SELECT Buyyear,sum(ZONGJINE) AS ZONGJINE, SUM(LIRUN) AS LIRUN,SUM(CHENGBEN) AS CHENGBEN ,SUM(YOUFEI) AS YOUFEI,SUM(SHULIANG) AS SHULIANG FROM RP_SALES_ALL_OrderNum "&tiaojian_y&" GROUP BY Buyyear order by Buyyear"
	Case else
		'SSQL="SELECT Success_time, sum(BuyPrice*ProdUnit) AS ZONGJINE,sum((BuyPrice-Product_Price)*ProdUnit) AS LIRUN,sum(Product_Price*ProdUnit) AS CHENGBEN ,sum(youfei) AS YOUFEI,sum(ProdUnit) AS SHULIANG  FROM RP_SALES_ALL WHERE Buyyearmonth='"&DatePart("yyyy",now())&right("0"&Month(now()),2)&"' GROUP BY Success_time order by Success_time"
		SSQL="SELECT sum(ZONGJINE) AS ZONGJINE, SUM(LIRUN) AS LIRUN,SUM(CHENGBEN) AS CHENGBEN ,SUM(YOUFEI) AS YOUFEI,SUM(SHULIANG) AS SHULIANG,BuyyearmonthDate FROM RP_SALES_ALL_OrderNum  WHERE Buyyearmonth='"&DatePart("yyyy",now())&right("0"&Month(now()),2)&"' GROUP BY BuyyearmonthDate order by BuyyearmonthDate"
		
End Select
response.write "<!--"&SSQL&"-->"
set rs=Server.Createobject("ADODB.RecordSet")
rs.open SSQL,conn,1,3
if rs.eof and rs.bof  then
	response.write "<tr  bgcolor=#FFFFFF><td height=25 colspan=10 align=center>��ʱû�����ݣ�</td></tr>"
else
'pages=50
'pagesnumber=pages
'totalrs=rs.RecordCount
'rs.pageSize = pages
'allPages = rs.pageCount
'page = Cint(Request("page"))
'If not isNumeric(page) then page=1
'if isEmpty(page) or int(page) <= 1 then
'page = 1
'elseif int(page) > allPages then
'page = allPages 
'end if
'rs.AbsolutePage = page
'	Total = 0
	Do While Not rs.eof
'	'Sum = FormatNumber(csng(rs("BuyPrice"))*csng(rs("ProdUnit")),2)-FormatNumber(csng(rs("Product_Price"))*csng(rs("ProdUnit")),2)
'	Sum = FormatNumber(Sum,2)
'	Total = Sum + Total
Dim Total_ZONGJINE,Total_SHULIANG,Total_LIRUN,Total_CHENGBEN,Total_YOUFEI
Total_ZONGJINE=Total_ZONGJINE+rs("ZONGJINE")
Total_SHULIANG=Total_SHULIANG+rs("SHULIANG")
Total_LIRUN=Total_LIRUN+rs("LIRUN")
Total_CHENGBEN=Total_CHENGBEN+rs("CHENGBEN")
Total_YOUFEI=Total_YOUFEI+rs("YOUFEI")
	%>
	<tr bgcolor=#FFFFFF>
	<td height=22 align="center"><%=rs(0)%></td>
	<td height="22" align="center" bgcolor="#CAE8F9"><%=formatnumber(rs("ZONGJINE"),2)%></td>
	<td height="22" align="center" bgcolor="#F0D0E7"><%=rs("SHULIANG")%></td>
	<td height="22" align=center bgcolor="#FCFAD1"><%=formatnumber(rs("LIRUN"),2)%></td>
	<td height="22" align=center bgcolor="#D6EBF3"><%=formatnumber(rs("CHENGBEN"),2)%></td>
	<td height="22" align=center><%=formatnumber(rs("YOUFEI"),2)%></td>
	</tr>
	<%
	'pages = pages - 1
	rs.movenext
if rs.eof then exit do
loop
%>
	<tr bgcolor=#FFFFFF>
	<td height=22 align="center">�ϼƣ�</td>
	<td height="22" align="center" bgcolor="#CAE8F9"><B><%=formatnumber(Total_ZONGJINE,2)%></B></td>
	<td height="22" align="center" bgcolor="#F0D0E7"><B><%=Total_SHULIANG%></B></td>
	<td height="22" align=center bgcolor="#FCFAD1"><B><%=formatnumber(Total_LIRUN,2)%></B></td>
	<td height="22" align=center bgcolor="#D6EBF3"><B><%=formatnumber(Total_CHENGBEN,2)%></B></td>
	<td height="22" align=center><B><%=formatnumber(Total_YOUFEI,2)%></B></td>
	</tr>
</table>
</BR>
<%
'call listpage()
end if
rs.close
set rs=nothing

FUNCTION listpage()
	response.write "¼�� <font color=red>"&totalrs&" �� </font> &nbsp; ÿҳ <font color=red> "&pagesnumber&" </font>&nbsp; "
if page = 1 then
	response.write "<font color=darkgray>��ҳ ǰҳ</font>"
else
	response.write "<a href=manage_order6.asp?StatusDefine="&Product_Factory&"&Recname2="&Recname2&"&ProdId="&ProdId&"&page=1>��ҳ</a> <a href=manage_order6.asp?StatusDefine="&Product_Factory&"&Recname2="&Recname2&"&ProdId="&ProdId&"&page="&page-1&">ǰҳ</a>"
end if

action="manage_order6.asp?StatusDefine="&Product_Factory&"&Recname2="&Recname2&"&ProdId="&ProdId&""				
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
	response.write " <a href=manage_order6.asp?StatusDefine="&Product_Factory&"&Recname2="&Recname2&"&ProdId="&ProdId&"&page="&page+1&">��ҳ</a> <a href=manage_order6.asp?StatusDefine="&Product_Factory&"&Recname2="&Recname2&"&ProdId="&ProdId&"&page="&allpages&">ĩҳ</a>"
end if
end FUNCTION
%>
</body>
</html>
