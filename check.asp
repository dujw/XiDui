<!--#include file="include/connect.asp" -->
<%
call aspsql()
'�ж��û��Ƿ��½
call checklogin()
Name   = request.cookies("buyok")("userid")
action = Trim(request("action"))
cateid = Trim(request("cateid"))
ProdId = Trim(request("ProdId"))
shul   = Trim(request("spec_"&Trim(request("ProdId"))))
chima  = Trim(request("cart_"&Trim(request("ProdId"))))

		set rsadd=server.createobject("adodb.recordset")
		sql="select * from dk501_border where UserId='"& Name &"' and OrderNum='"& Name &"'"
		rsadd.open sql,conn,1,1
		if (rsadd.bof and rsadd.eof) then
			response.redirect "error/error.asp?error=007"
			response.end
		end if
if action = "del" then
	if cateid <> "" then
		conn.execute "delete from dk501_border where id="&cateid
			response.redirect "check.asp"
			response.End
		else
			response.write "<script language='javascript'>"
			response.write "alert('����ʧ�ܣ�û������ɾ�������ݡ�');"
			response.write "location.href='check.asp';"			
			response.write "</script>"
	end if
end if

'��ʼ�༭���������뵫�ǲ��ܱ༭��
If Request("edit") = "ok" Then
		set rs_Product =Server.Createobject("ADODB.RecordSet")
		sql_Product="select ProdId,ProdUnit,xinhao from dk501_border where  UserId='"& Name &"' and OrderNum='"& Name &"'"
		rs_Product.Open sql_Product,conn,1,3
		if not rs_Product.eof then
			if Trim(request("spec_"&rs_Product("ProdId")&"_"&rs_Product("xinhao")))="" then
				response.write "<script language='javascript'>"
				response.write "alert('��ʾ������������������Ҫ��������');"
				response.write "location.href='check.asp';"			
				response.write "</script>"
			else
				do while not rs_Product.eof
					'�������� ProdUnit
				conn.execute("update dk501_border set ProdUnit= '"&Cint(Trim(request("spec_"&rs_Product("ProdId")&"_"&rs_Product("xinhao")))) &"' where ProdId='"& rs_Product("ProdId") &"' and UserId='"& Name &"' and OrderNum='"& Name &"' and xinhao = '"& Trim(request("cart_"&rs_Product("ProdId")&"_"&rs_Product("xinhao"))) &"'")
					rs_Product.movenext
				loop
				end if	
				rs_Product.close
				set rs_Product=nothing					
	end if
end if
%>
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
<div id="div_Public">
  <table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"></td>
    </tr>
    <tr>
      <td align="center" class="b-top b-left b-right b-bottom" style="padding:6px 6px 6px 6px;">
	  <table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
        <td height="65" align="left" background="../image/order/news2.jpg" class="b-left b-right">&nbsp;&nbsp;<img src="image/order/user1.gif" width="11" height="11" /> ��ǰλ�ã�<a href="index.asp">��ҳ</a> ��<a href="shop.asp">��������</a> ���ҵĶ���</td>
      </tr>
      <tr>
        <td height="28" align="left" valign="middle" background="../image/order/bucou1.jpg" class="b-left b-right">&nbsp;&nbsp;&nbsp;<font color="#FFFFFF">�����ύ���1.ѡ���Ʒ������Ʒ 2.��дҪ�ĳ��루��������������Ʒ��3.��д��ϵ��ʽ�Լ����� 4.�ɹ��ύ����</font></td>
  </tr>
      <tr>
        <td height="89" align="left" valign="top" class="b-left b-right b-bottom" style="padding-top:10px; padding-right:10px; padding-bottom:10px; padding-left:10px;">
          <table border="0" align="center" width="870" cellpadding="0" cellspacing="0">
            <form action="check.asp" method="post" name="check">
              <tr align="center" bgcolor="#EEEEEE">
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#999999">
                    <tr bgcolor="#f0f0f0">
                      <td height="23" colspan="2" align="center">�� Ʒ �� ��</td>
                      <td width="65" height="23" align="center">�г���</td>
                      <td width="65" height="23" align="center">����۸�</td>
                      <td width="65" align="center">��Ʒ����</td>
                      <td width="55" align="center">�� ��</td>
                      <td width="120" height="23" align="center" nowrap="nowrap">����</td>
                      <td width="94" height="23" align="center">�༭</td>
                    </tr>
					<%
					Set rs=conn.execute("select A.id,A.ProdId,A.ProdUnit,A.BuyPrice,A.xinhao,B.ProdName,B.ImgPrev,B.PriceOrigin,B.zhongliang,B.jifenzhi from dk501_border A,dk501_product B where A.ProdId = B.ProdId and  A.UserId='"& Name &"' and A.OrderNum='"& Name &"' order by A.ProdId")
					if rs.eof and rs.bof then
						response.redirect "error/error.asp?error=007"
						response.end
					else
						Total = 0
						TotalA= 0
						TotalB= 0
						do while not rs.eof
						Sum = formatnum(csng(rs("BuyPrice"))*rs("ProdUnit"),2)
						Sum = FormatNumber(Sum,2)
						SumA = formatnum(csng(rs("zhongliang"))*rs("ProdUnit"),2)
						SumB = formatnum(csng(rs("jifenzhi"))*rs("ProdUnit"),2)
						Total = Sum + Total
						TotalA = SumA + TotalA
						TotalB = SumB + TotalB
					%>
                    <tr bgcolor="#FFFFFF">
                      <td width="53" height="52" align="center" valign="middle"><img src="<%=rs("ImgPrev")%>" width="50" height="50"></td>
                      <td width="355" align="left" valign="middle"><a href='list.asp?ProdId=<%=rs("ProdId")%>'><%=rs("ProdName")%></a></td>
                      <td align="center" class="shich_cont"><%=FormatNum(Trim(rs("PriceOrigin")),2)%></td>
                      <td align="center" nowrap="nowrap"><%=FormatNum(Trim(rs("BuyPrice")),2)%>Ԫ</td>
                      <td align="center" nowrap="nowrap"><%=rs("zhongliang")%>��</td>
                      <td align="center"><input type="text" name="spec_<%=rs("ProdId")%>_<%=rs("xinhao")%>" class="uset_input" size="4" maxlength="4"  value="<%=rs("ProdUnit")%>" /></td>
                      <td align="center"><input type="text" name="cart_<%=rs("ProdId")%>" class="uset_input" size="10" value=<%=rs("xinhao")%> disabled="disabled">
                      <input type="hidden" name="cart_<%=rs("ProdId")%>_<%=rs("xinhao")%>" class="uset_input" size="10" value=<%=rs("xinhao")%>>
                      <input type="hidden" name="ProdId" class="uset_input" size="10" value=<%=rs("ProdId")%>></td>
                      <td align="center" nowrap="nowrap">
					  <input type="submit" name="order" value="ȷ��"><br>
					 <a href="check.asp?cateid=<%=rs("id")%>&action=del">ɾ��</a>					 </td>
                  </tr>
				  <%
				  rs.movenext
				  loop
				  rs.close
				  set rs=nothing
				  end if 
				  %>
                </table></td>
              </tr>
              <tr align="left">
                <td bgcolor="#ffffff" height="35">������ܼƣ�<font color="#FF0000">��<%=formatnumber(total,2)%> </font> ���أ�<font color="#FF0000"><%=formatnumber(totalA,2)%>��</font> ����û��֣�<%=totalB%></td>
              </tr>
              <tr>
                <td colspan="6" bgcolor="#ffffff" height="40" align="right">
				<a href=index.asp><img src="image/order/goonshopping.jpg" width="92" height="29" border="0"></a>
				<a href=payment.asp><img src=image/order/pay.jpg width="63" height="29" border=0></a>
				<input type="hidden" name="edit" value="ok">				</td>
              </tr>
			  </form>
        </table></td>
      </tr>
</table>	  </td>
    </tr>
  </table>
</div>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>
