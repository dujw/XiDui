<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../include/connect.asp"-->
<%
call aspsql()
id=trim(request("prodid"))
if id="" then response.redirect "main.asp"
	Set rsprod=conn.execute ("select ProdId,ProdName,ImgPrev2,Prod1,LarCode,MidCode,ClickTimes,PriceOrigin,PriceList,PriceList1,PriceList2,PriceList3,PriceList4,PriceList5,PriceList6,Prod1,zhongliang,jifenzhi,adddate,MemoSpec,ProdDisc from dk501_product where  ProdId='"&id&"'")
	if (rsprod.bof and rsprod.eof) then
		response.redirect "index.asp"
	else
		Prodname=rsprod("Prodname")	
		LarCod=rsprod("LarCode")	
		Midcod=rsprod("MidCode")	
		prodid=rsprod("prodid")
		conn.execute "UPDATE dk501_product SET ClickTimes ="&rsprod("ClickTimes")+1&" WHERE ProdId ='"&id&"'"
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="../image/web_css/login.css" rel="stylesheet" type="text/css" />
</head>
<body>
<!--#include file="../include/top.asp"-->
<div id="div_Public">
  <table width="950" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="40" colspan="3" background="../image/product/list.jpg" style="padding-left:200px;">
	  <a href="/index.asp" class="text">��ҳ</a>&nbsp;|&nbsp;
	  <%
	  SubLar=trim(request("LarCode"))
	  set rs_da=server.createobject("adodb.recordset") 
	  sql_da="select LarCode from dk501_productclass where MidSeq = 1 order by LarSeq asc" 
	  rs_da.Open sql_da,conn,1,1
	  if rs_da.bof and rs_da.eof then
		  response.write "��ʱû�з���"
	  else
	  	  Do While Not rs_da.eof	
	  %>
	  <a href="/product/product.asp?LarCode=<%=Server.URLEncode(Trim(rs_da("LarCode")))%>" class=text><%=rs_da("LarCode")%></a>&nbsp;|&nbsp;<%
	  rs_da.movenext
		  if rs_da.eof then exit do
		loop		
	end if
	rs_da.close
	set rs_da=nothing
	  %>	  </td>
    </tr>
    <tr>
      <td colspan="3"><img src="../adv/list.jpg" width="950" height="90" /></td>
    </tr>
    <tr>
      <td height="30" colspan="3" class="b-bottom"><span class="b-bottom"><img src="product_1.png" width="16" height="16" /> Ŀǰλ��  <a href="main.asp">��ҳ</a> > 
	<a href=product.asp?LarCode=<%=Trim(rsprod("LarCode"))%>><%=Trim(rsprod("LarCode"))%></a>  > 
	<a href=product.asp?LarCode=<%=Trim(rsprod("LarCode"))%>&MidCode=<%=Trim(rsprod("MidCode"))%>><%=Trim(rsprod("MidCode"))%></a> > 
	<%=Trim(rsprod("Prodname"))%></span></td>
    </tr>
    <tr>
      <td height="5" colspan="3"></td>
    </tr>
    <tr>
      <td width="302" rowspan="2" align="center" class="b-top b-left b-right b-bottom"><span >
<img src='/<%=rsprod("ImgPrev2")%>' alt='<%=rsprod("ProdName")%>' width="300" height="300"  border="0" /></span></td>
      <td width="200" rowspan="2" align="center" valign="top" style="padding:6px 6px 6px 6px;"><table width="390" height="284" border="0" align="left" cellpadding="0" cellspacing="0">
        <tr>
          <td height="30" align="left" class="Brown"><%=lleft(rsprod("ProdName"),80)%></td>
        </tr>
        <tr>
          <td height="25" align="left"><span class="Brown">�ϼ�ʱ�䣺<%=rsprod("adddate")%></span></td>
        </tr>
        <tr>
          <td height="25" align="left">��Ʒ������<%=rsprod("zhongliang")%>g</td>
        </tr>
        <tr>
          <td height="25" align="left">ÿ����һ������ <span class="Brown"><%=rsprod("jifenzhi")%> </span>�����֣� </td>
        </tr>
		<form action="shop.asp?ProdId=<%=rsprod("ProdId")%>" method='post'>
        <tr>
          <td height="131" align="left" valign="top" bgcolor="#FFF3D9"  class="b-top b-left b-right b-bottom" style="padding:6px 6px 6px 6px;"> 
		  �г��۸� ��<%=FormatNum(rsprod("PriceOrigin"),2)%> <br />
		  ����ۼۣ� ��<%=FormatNum(rsprod("PriceList"),2)%>
		  <% if request.cookies("buyok")("userid")<> "" then 
			Set rstxt = conn.Execute("select usertype from dk501_user where userid='"&request.cookies("buyok")("userid")&"'") 
			if not rstxt.eof then
			usertype=Trim(rstxt("usertype"))
			end if
			set rstxt=nothing
				 if usertype="1" then
					response.write "<br><font color=red>"&type1&"����"&FormatNum(rsprod("PriceList1"),2)&"<input name=Type_"&rsprod("ProdId")&" type=hidden value="&rsprod("PriceList1")&"></font><br>"
				 elseif usertype="2" then
					response.write "<br><font color=red>"&type2&"����"&FormatNum(rsprod("PriceList2"),2)&"<input name=Type_"&rsprod("ProdId")&" type=hidden value="&rsprod("PriceList2")&"></font><br>"
				 elseif usertype="3" then
					response.write "<br><font color=red>"&type3&"����"&FormatNum(rsprod("PriceList3"),2)&"<input name=Type_"&rsprod("ProdId")&" type=hidden value="&rsprod("PriceList3")&"></font><br>"
				 elseif usertype="4" then
					response.write "<br><font color=red>"&type4&"����"&FormatNum(rsprod("PriceList4"),2)&"<input name=Type_"&rsprod("ProdId")&" type=hidden value="&rsprod("PriceList4")&"></font><br>"
				 elseif usertype="5" then
					response.write "<br><font color=red>"&type5&"����"&FormatNum(rsprod("PriceList5"),2)&"<input name=Type_"&rsprod("ProdId")&" type=hidden value="&rsprod("PriceList5")&"></font><br>"
				 elseif usertype="6" then
					response.write "<br><font color=red>"&type6&"����"&FormatNum(rsprod("PriceList6"),2)&"<input name=Type_"&rsprod("ProdId")&" type=hidden value="&rsprod("PriceList6")&"></font><br>"
				 end if	  
			end if%>
		  <br />
		  ���г��룺
		  <%if rsprod("Prod1")<>"" then
		  size=trim(rsprod("Prod1"))
		  myarr=split(size,"/")
		  for i=0 to ubound(myarr) %>
		  </option><input name="spec_<%=Trim(request("ProdId"))%>" type="radio" value="<%=trim(myarr(i))%>" /><%=trim(myarr(i))%>
          <%next%>
		  <%else
		  response.write "�޳�����Ϣ" 
		  end if%>		  <br />
		  ����������<input name="cart_<%=Trim(request("ProdId"))%>" type="text" value="1" onKeyUp="value=value.replace(/[^\d]/g,'') " size="5">
		  <a href="onlineflow.asp">
		  <input name="image" type="image" src="dk501_32.jpg" />
		  </a><br />
		  <br />
		  <a href="/Contact/Contact.asp?id=1">��������</a> <a href="/Contact/Contact.asp?id=2">&nbsp;&nbsp;����˵��</a> &nbsp;&nbsp;<a href="/Contact/Contact.asp?id=3">��ͨ���</a>&nbsp; <a href="/Contact/Contact.asp?id=4">����ϸ��</a> <a href="/Contact/Contact.asp?id=6">&nbsp;&nbsp;��ϵ����</a> <br />
		  <a href="/Contact/Contact.asp?id=7">��λ���</a> <a href="/Contact/Contact.asp?id=8">&nbsp;&nbsp;������֪</a> <a href="/Contact/Contact.asp?id=9">&nbsp;&nbsp;����˿�</a> &nbsp;<a href="/Contact/Contact.asp?id=10">������ѯ</a>&nbsp;&nbsp; <a href="/Contact/Contact.asp?id=12">��������</a></td>
        </tr>
		</form>
      </table></td>
      <td width="250" height="30" class="b-top b-left b-right">&nbsp;���߿ͷ�</td>
    </tr>
    <tr>
      <td height="270" valign="top" class="b-top b-left b-right b-bottom" style="padding-top:10px; padding-left:10px; ">
	  ���߿ͷ�1��<a target=blank href=tencent://message/?uin=821653210&Site=����Сѩ&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:770841592:4 alt="����Сѩ">����ͷ�Сѩ&nbsp;</a><br>
	  <br>
	  ���߿ͷ�2��<a target=blank href=tencent://message/?uin=821653210&Site=����Сѩ&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:770841592:4 alt="����Сѩ">����ͷ�Сѩ&nbsp;</a><br><br>
	  �ۺ�ͷ�1��<a target=blank href=tencent://message/?uin=821653210&Site=����Сѩ&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:770841592:4 alt="����Сѩ">����ͷ�Сѩ&nbsp;</a><br><br>
	  �ۺ�ͷ�2��<a target=blank href=tencent://message/?uin=821653210&Site=����Сѩ&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:770841592:4 alt="����Сѩ">����ͷ�Сѩ&nbsp;</a><br><br>
	  �ƹ�ͷ�1��<a target=blank href=tencent://message/?uin=821653210&Site=����Сѩ&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:770841592:4 alt="����Сѩ">����ͷ�Сѩ&nbsp;</a><br><br>
	  ���˿ͷ�1��<a target=blank href=tencent://message/?uin=821653210&Site=����Сѩ&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:770841592:4 alt="����Сѩ">����ͷ�Сѩ&nbsp;</a><br><br>
	  </td>
    </tr>
    
	<tr>
      <td colspan="3" height="5"></td>
    </tr>
	<tr>
      <td colspan="3"><img src="../image/product/aa.jpg" width="103" height="29" /></td>
    </tr>
	<tr>
      <td colspan="3" align="center" class="b-top b-left b-right b-bottom" style="padding-top:10px; padding-bottom:10px; "><span class="b-all" style="padding-top:10px; padding-bottom:10px; "><img src="/<%=rsprod("MemoSpec")%>" /></span></td>
    </tr>
  </table>
  <% end if %>
</div>
<!--#include file="../include/login/bootom.asp"-->
</body>
</html>
<% call CloseConn() %>
