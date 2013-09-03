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
	  <a href="/index.asp" class="text">首页</a>&nbsp;|&nbsp;
	  <%
	  SubLar=trim(request("LarCode"))
	  set rs_da=server.createobject("adodb.recordset") 
	  sql_da="select LarCode from dk501_productclass where MidSeq = 1 order by LarSeq asc" 
	  rs_da.Open sql_da,conn,1,1
	  if rs_da.bof and rs_da.eof then
		  response.write "暂时没有分类"
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
      <td height="30" colspan="3" class="b-bottom"><span class="b-bottom"><img src="product_1.png" width="16" height="16" /> 目前位置  <a href="main.asp">首页</a> > 
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
          <td height="25" align="left"><span class="Brown">上架时间：<%=rsprod("adddate")%></span></td>
        </tr>
        <tr>
          <td height="25" align="left">商品重量：<%=rsprod("zhongliang")%>g</td>
        </tr>
        <tr>
          <td height="25" align="left">每购买一件赠送 <span class="Brown"><%=rsprod("jifenzhi")%> </span>个积分！ </td>
        </tr>
		<form action="shop.asp?ProdId=<%=rsprod("ProdId")%>" method='post'>
        <tr>
          <td height="131" align="left" valign="top" bgcolor="#FFF3D9"  class="b-top b-left b-right b-bottom" style="padding:6px 6px 6px 6px;"> 
		  市场价格： ￥<%=FormatNum(rsprod("PriceOrigin"),2)%> <br />
		  最低售价： ￥<%=FormatNum(rsprod("PriceList"),2)%>
		  <% if request.cookies("buyok")("userid")<> "" then 
			Set rstxt = conn.Execute("select usertype from dk501_user where userid='"&request.cookies("buyok")("userid")&"'") 
			if not rstxt.eof then
			usertype=Trim(rstxt("usertype"))
			end if
			set rstxt=nothing
				 if usertype="1" then
					response.write "<br><font color=red>"&type1&"：￥"&FormatNum(rsprod("PriceList1"),2)&"<input name=Type_"&rsprod("ProdId")&" type=hidden value="&rsprod("PriceList1")&"></font><br>"
				 elseif usertype="2" then
					response.write "<br><font color=red>"&type2&"：￥"&FormatNum(rsprod("PriceList2"),2)&"<input name=Type_"&rsprod("ProdId")&" type=hidden value="&rsprod("PriceList2")&"></font><br>"
				 elseif usertype="3" then
					response.write "<br><font color=red>"&type3&"：￥"&FormatNum(rsprod("PriceList3"),2)&"<input name=Type_"&rsprod("ProdId")&" type=hidden value="&rsprod("PriceList3")&"></font><br>"
				 elseif usertype="4" then
					response.write "<br><font color=red>"&type4&"：￥"&FormatNum(rsprod("PriceList4"),2)&"<input name=Type_"&rsprod("ProdId")&" type=hidden value="&rsprod("PriceList4")&"></font><br>"
				 elseif usertype="5" then
					response.write "<br><font color=red>"&type5&"：￥"&FormatNum(rsprod("PriceList5"),2)&"<input name=Type_"&rsprod("ProdId")&" type=hidden value="&rsprod("PriceList5")&"></font><br>"
				 elseif usertype="6" then
					response.write "<br><font color=red>"&type6&"：￥"&FormatNum(rsprod("PriceList6"),2)&"<input name=Type_"&rsprod("ProdId")&" type=hidden value="&rsprod("PriceList6")&"></font><br>"
				 end if	  
			end if%>
		  <br />
		  现有尺码：
		  <%if rsprod("Prod1")<>"" then
		  size=trim(rsprod("Prod1"))
		  myarr=split(size,"/")
		  for i=0 to ubound(myarr) %>
		  </option><input name="spec_<%=Trim(request("ProdId"))%>" type="radio" value="<%=trim(myarr(i))%>" /><%=trim(myarr(i))%>
          <%next%>
		  <%else
		  response.write "无尺码信息" 
		  end if%>		  <br />
		  购买数量：<input name="cart_<%=Trim(request("ProdId"))%>" type="text" value="1" onKeyUp="value=value.replace(/[^\d]/g,'') " size="5">
		  <a href="onlineflow.asp">
		  <input name="image" type="image" src="dk501_32.jpg" />
		  </a><br />
		  <br />
		  <a href="/Contact/Contact.asp?id=1">关于我们</a> <a href="/Contact/Contact.asp?id=2">&nbsp;&nbsp;付款说明</a> &nbsp;&nbsp;<a href="/Contact/Contact.asp?id=3">普通快递</a>&nbsp; <a href="/Contact/Contact.asp?id=4">配送细节</a> <a href="/Contact/Contact.asp?id=6">&nbsp;&nbsp;联系我们</a> <br />
		  <a href="/Contact/Contact.asp?id=7">如何换货</a> <a href="/Contact/Contact.asp?id=8">&nbsp;&nbsp;代理须知</a> <a href="/Contact/Contact.asp?id=9">&nbsp;&nbsp;如何退款</a> &nbsp;<a href="/Contact/Contact.asp?id=10">批发咨询</a>&nbsp;&nbsp; <a href="/Contact/Contact.asp?id=12">帮助中心</a></td>
        </tr>
		</form>
      </table></td>
      <td width="250" height="30" class="b-top b-left b-right">&nbsp;在线客服</td>
    </tr>
    <tr>
      <td height="270" valign="top" class="b-top b-left b-right b-bottom" style="padding-top:10px; padding-left:10px; ">
	  在线客服1：<a target=blank href=tencent://message/?uin=821653210&Site=代理小雪&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:770841592:4 alt="代理小雪">代理客服小雪&nbsp;</a><br>
	  <br>
	  在线客服2：<a target=blank href=tencent://message/?uin=821653210&Site=代理小雪&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:770841592:4 alt="代理小雪">代理客服小雪&nbsp;</a><br><br>
	  售后客服1：<a target=blank href=tencent://message/?uin=821653210&Site=代理小雪&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:770841592:4 alt="代理小雪">代理客服小雪&nbsp;</a><br><br>
	  售后客服2：<a target=blank href=tencent://message/?uin=821653210&Site=代理小雪&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:770841592:4 alt="代理小雪">代理客服小雪&nbsp;</a><br><br>
	  推广客服1：<a target=blank href=tencent://message/?uin=821653210&Site=代理小雪&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:770841592:4 alt="代理小雪">代理客服小雪&nbsp;</a><br><br>
	  查账客服1：<a target=blank href=tencent://message/?uin=821653210&Site=代理小雪&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:770841592:4 alt="代理小雪">代理客服小雪&nbsp;</a><br><br>
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
