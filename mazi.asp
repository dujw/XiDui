<!--#include file="include/connect.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
<% 
SubLar=trim(request("LarCode")) 
SubCode=trim(request("MidCode"))
	function cutstr(tempstr,tempwid)
		if len(tempstr)>tempwid then
			cutstr=left(tempstr,tempwid)&"..."
		else
			cutstr=tempstr
		end if
	end function
%>
</head>
<body>
<!--#include file="include/web_top.asp" -->
<table width="950" height="36" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="900" height="30" valign="middle" class="b-bottom"><A name=bsc></A><span> 
	目前位置 <a href="main.asp">首页</a> > <font color="#FF0000"><%=Request("title")%></font></td>
  </tr>
  <tr>
    <td height="10" ></td>
  </tr>
</table>
<TABLE cellSpacing=0 cellPadding=0 width=950 
                                align=center border=0>
  <TBODY>
    <TR>
      <TD width="371" height=187 align="center"><% 
		Quantity=Trim(Request("Quantity"))
		set rs=server.CreateObject("adodb.recordset")		
		sql="select ProdId,ImgPrev,ProdName,Prod1,PriceOrigin,PriceList,PriceList1,PriceList2,PriceList3,PriceList4,PriceList5,PriceList6,AddDate from dk501_product where Quantity='"&Quantity&"' and online = 2 order by AddDate desc"
		rs.Open sql,conn,1,1	
		 %>
          <table width="900" height="163"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr align="center">
              <td width="550" height="120"><!--#include file="page3.asp"-->
                  <table width="160" height="170" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#FF9900">
                    <%
          j=1
          for i=1 to x
          if (j mod 6)=1 then
              response.write "<tr>"
          end if
        %>
                    <td width="240" align="center" valign="top" whith="170"><table width="160" height="150" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="193" height="170" colspan="3" align="center" ><table width="152" border="0" cellpadding="0" cellspacing="0">
                          <tr >
                            <td width="161" height="160" colspan="2" align="center" class="b-top b-left b-right b-bottom"><A href="list.asp?ProdId=<%=rs("ProdId")%>" target=_blank><img src="<%=rs("ImgPrev")%>" width="150" height="150" border="0"></A></td>
                          </tr>
                          <tr>
                            <td height="4" colspan="2" align="center"></td>
                          </tr>
                          <tr>
                            <td height="30" colspan="2" align="left"><A href="list.asp?ProdId=<%=rs("ProdId")%>" target=_blank><%=lleft(rs("ProdName"),20)%></a> <BR>
                                  <span class="product_font">市场价：</span>￥<%=formatnumber(rs("PriceOrigin"),2)%><BR>
                                <SPAN class=hy_cont>最低售价：￥<font color=#FF0000><%=formatnumber(rs("PriceList"),2)%></font></SPAN><br>
<%
		if request("Quantity")=0 then 
			response.Write "下架时间："
		else 
			response.Write "上架时间："
		end if
		%>
        <%=split(rs("AddDate")," ")(0)%></td>
                          </tr>
                          <tr>
                            <td height="10" colspan="2" align="center"></td>
                          </tr>
                        </table></td>
                      </tr>
                    </table></td>
                    <%  
          if (j mod 6)=0 then
		     response.write "</tr>"
          end if
	      j=j+1
		  rs.moveNext
          next
       %>
                </table></td>
            </tr>
            <tr align="center">
              <td valign="middle" nowrap><!--#include file="page4.asp"--></td>
            </tr>
        </table></TD>
    </TR>
  </TBODY>
</TABLE>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>
