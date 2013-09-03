<!--#include file="include/connect.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
</head>
<body>
<!--#include file="include/web_top.asp" -->
<table width="950" height="400" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="710" rowspan="3"><!--#include file="include/adv.asp" --></td>
    <td width="10" rowspan="3">&nbsp;</td>
    <td width="230"><a href="<%=index_right_url_1%>" target="_blank"><img src="<%=index_right_1%>" border=0></a></td>
  </tr>
  <tr>
    <td valign="middle"><a href="<%=index_right_url_2%>" target="_blank"><img src="<%=index_right_2%>" border=0></a></td>
  </tr>
  <tr>
    <td valign="bottom"><a href="<%=index_right_url_3%>" target="_blank"><img src="<%=index_right_3%>" border=0></a></td>
  </tr>
</table>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3" style="border-bottom:solid 2px #A10000;"><img src="images/index/newProduct.jpg" width="264" height="55"></td>
  </tr>
  <tr>
    <td colspan="3" height="6"></td>
  </tr>
  <tr>
    <td width="277"><a href=<%=index_adv_url_1%> target="_blank"><img src="<%=index_adv_1%>" border=0></a></td>
    <td width="436" align="center"><a href=<%=index_adv_url_2%> target="_blank"><img src="<%=index_adv_2%>" border=0></a></td>
    <td width="247" align="right" valign="top"><a href=<%=index_adv_url_3%> target="_blank"><img src="<%=index_adv_3%>" border=0></a></td>
  </tr>
  <tr>
    <td colspan="3" height="6"></td>
  </tr>
  <tr>
    <td colspan="3" style="border-bottom:solid 2px #A10000;"><img src="images/index/men.jpg" width="264" height="55"></td>
  </tr>
  <tr>
    <td colspan="3" height="6"></td>
  </tr>
  <tr>
    <td colspan="3" height="6" class="b-top b-left b-right b-bottom" style="padding-top:4px; padding-right:4px; padding-bottom:4px; padding-left:4px;"><table width="942" border="0" align="center" cellpadding="0" cellspacing="0" >
      <tr>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.open "select top 30 ProdId,ImgPrev,ProdName,Prod1,PriceOrigin,PriceList,AddDate from dk501_product where online=2 and Quantity =1   order by AddDate desc ",conn,1,1
		  if rs.eof and rs.bof then
		  response.write "<center><br><font color=red size=2>对不起，暂无此类商品！</font></font>"
		  'response.End
		  else
		  %>
        <%
		if not rs.eof then
		i=1
		do while not rs.eof%>
        <td height="110" valign="top" style="padding-top:4px;"><table border="0" cellspacing="0" cellpadding="0" valign="middle" width="140" >
            <tr>
              <td width="116" height="116" align="center" valign="middle" background="images/bg/1.jpg" class="b-top b-left b-right b-bottom"><a href='list.asp?ProdId=<%=rs("ProdId")%>' target="_blank"><img  border="0" alt="<%=rs("ProdName")%>" src="<%=rs("ImgPrev")%>" width="180" /></a></td>
            </tr>
            <tr>
              <td height="16" align="left"><a href='list.asp?ProdId=<%=rs("ProdId")%>' target="_blank"><%=lleft(rs("ProdName"),14)%></a></td>
            </tr>
            <tr>
              <td height="16" align="left"><font color="#DB3578">尺码：<%=lleft(Trim(rs("Prod1")),16)%></font></td>
            </tr>
            <tr>
              <td height="16" align="left">市场价：￥<%=formatnumber(rs("PriceOrigin"),2)%></td>
            </tr>
        </table></td>
        <%if i mod 5 = 0 then%>
      </tr>
      <%end if
	    rs.movenext
         i=i+1
    	 loop
		rs.close
		set rs=nothing
		end if
		end if 
	%>
    </table></td>
  </tr>
  <tr>
    <td colspan="3" height="6"></td>
  </tr>
</table>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>