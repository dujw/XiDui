<!--#include file="include/connect.asp" -->
<%
call aspsql()
id=trim(request("prodid"))
if id="" then response.redirect "main.asp"
	Set rsprod=conn.execute ("select ProdId,ProdName,ImgPrev2,ImgPrev,Prod1,MemoSpec_PSD,MemoSpec_AD,LarCode,MidCode,ClickTimes,PriceOrigin,PriceList,PriceList1,PriceList2,PriceList3,PriceList4,PriceList5,PriceList6,Prod1,zhongliang,jifenzhi,adddate,MemoSpec,ProdDisc from dk501_product where  ProdId='"&id&"'")
	if (rsprod.bof and rsprod.eof) then
		response.redirect "index.asp"
	else
		Prodname=rsprod("Prodname")	
		LarCod=rsprod("LarCode")	
		Midcod=rsprod("MidCode")	
		prodid=rsprod("prodid")
		conn.execute "UPDATE dk501_product SET ClickTimes ="&rsprod("ClickTimes")+1&" WHERE ProdId ='"&id&"'"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
<SCRIPT src="plus/discuz/js/common.js" type=text/javascript></SCRIPT>
<SCRIPT src="plus/discuz/js/viewthread.js" type=text/javascript></SCRIPT>
<link href="plus/discuz/css/style_1_viewthread.css" rel="stylesheet" />
</head>
<BODY id=viewthread onkeydown="if(event.keyCode==27) return false;">
<!--#include file="include/web_top.asp" -->
<DIV id=append_parent></DIV>
<div id="div_Public">
  <table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="30" colspan="3" class="b-bottom"><img src="product/product_1.png" width="16" height="16" /> 目前位置  <a href="main.asp">首页</a> > 
	<a href=product.asp?LarCode=<%=Trim(rsprod("LarCode"))%>><%=Trim(rsprod("LarCode"))%></a>  > 
	<a href=product.asp?LarCode=<%=Trim(rsprod("LarCode"))%>&MidCode=<%=Trim(rsprod("MidCode"))%>><%=Trim(rsprod("MidCode"))%></a> > 
	<%=Trim(rsprod("Prodname"))%></td>
    </tr>
    <tr>
      <td height="5" colspan="3"></td>
    </tr>
    <tr>
      <td align="center" class="b-top b-left b-right b-bottom">请点击图片后看大图</td>

      <td width="200" rowspan="2" align="center" valign="top" style="padding:6px 6px 0px 6px;"><table width="390" border="0" align="left" cellpadding="0" cellspacing="0">
        <tr>
          <td height="22" align="left" class="Brown"><%=lleft(rsprod("ProdName"),80)%></td>
        </tr>
        <tr>
          <td height="22" align="left"><span class="Brown">上架时间：<%=rsprod("adddate")%></span></td>
        </tr>
        <tr>
          <td height="22" align="left">商品重量：<%=rsprod("zhongliang")%>g</td>
        </tr>
        <tr>
          <td height="22" align="left">每购买一件赠送 <span class="Brown"><%=rsprod("jifenzhi")%> </span>个积分！ </td>
        </tr>
		<tr>
          <td height="22" align="left">&nbsp; </td>
        </tr>
		<form action="shop.asp?ProdId=<%=rsprod("ProdId")%>" method='post'>
        <tr>
          <td align="left" valign="top" bgcolor="#FFF3D9"  class="b-top b-left b-right b-bottom" style="padding:6px 6px 0px 6px;" height="200"> 
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
		  <input name="image2" type="image" src="dk501_32.jpg" />
		  </a><br />
		  <br />
		  <a href="Contact.asp?id=1">关于我们</a> <a href="Contact.asp?id=2">&nbsp;&nbsp;付款说明</a> &nbsp;&nbsp;<a href="Contact.asp?id=3">普通快递</a>&nbsp; <a href="Contact.asp?id=4">配送细节</a> <a href="Contact.asp?id=6">&nbsp;&nbsp;联系我们</a> <br />
		  <a href="Contact.asp?id=7">如何换货</a> <a href="Contact.asp?id=8">&nbsp;&nbsp;代理须知</a> <a href="Contact.asp?id=9">&nbsp;&nbsp;如何退款</a> &nbsp;<a href="Contact.asp?id=10">批发咨询</a>&nbsp;&nbsp; <a href="Contact.asp?id=12">帮助中心</a>
          </td>
        </tr>
		</form>
      </table></td>
      <td width="250" height="20" class="b-top b-left b-right">&nbsp;在线客服</td>
    </tr>
    <tr>
      <td align="center" class="b-top b-left b-right b-bottom"><A href="javascript:;"><img onClick="zoom(this, '<%=rsprod("ImgPrev2")%>')"  src='<%=rsprod("ImgPrev2")%>' alt='<%=rsprod("ProdName")%>'  border="0"/></a></td>
      <td height="225" valign="top" class="b-top b-left b-right b-bottom" style="padding-top:10px; padding-left:10px; ">
	   在线客服1：<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1504490396&site=裁难小刘&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:1504490396:45" alt="点击这里给我发消息" title="点击这里给我发消息">裁男小刘</a><br><br>
	
 在线客服2：<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2262462307&site=裁难小周&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2262462307:45" alt="点击这里给我发消息" title="点击这里给我发消息">裁男小周</a><br><br>
	
	
 在线客服3：<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2404470991&site=裁难小杨&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2404470991:45" alt="点击这里给我发消息" title="点击这里给我发消息">裁男小杨</a><br><br>
	
 在线客服4：<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2486488687&site=裁难小许&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2486488687:45" alt="点击这里给我发消息" title="点击这里给我发消息">裁男小许</a><br><br>

	 	  </td>
    </tr>
    
	<tr>
      <td colspan="3" height="5"></td>
    </tr>
	<tr>
      <td colspan="3"><img src="image/product/aa.jpg" width="103" height="29" /></td>
    </tr>
	<tr>
      <td colspan="3" align="center" class="b-top b-left b-right b-bottom" style="padding-top:10px; padding-bottom:10px; ">
      <%if rsprod("MemoSpec_AD")<>"" then%>
      	<img src="<%=rsprod("MemoSpec_AD")%>"/> <br/>
      <%end if%>
      <%if request.cookies("buyok")("userid")<>"" and rsprod("MemoSpec_PSD")<>"" then%>
        <span style="color:#F00; font-weight:bold;"><IMG src="images/ps.gif"> <a href="<%=rsprod("MemoSpec_PSD")%>" target="_blank">商品介绍PSD源文件下载</a>(请右键另存为或者使用迅雷下载) </span><br/><br/>
      <%end if%>  
      <img src="<%=rsprod("MemoSpec")%>" /></td>
    </tr>
  </table>
  <% end if %>
</div>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>
