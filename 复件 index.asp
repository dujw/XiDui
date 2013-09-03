<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/connect.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="description" content="<%=sitedescription%>">
<meta name="keywords" content="<%=sitekeywords%>">
<link href="image/web_css/login.css" rel="stylesheet" type="text/css" />
</head>
<body>
<%
	if kaiguan = "0" then 
		response.Write "<br><center>"&guanbi&"</font>"
		response.End()
	end if 
%>
<table width="950" height="78" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" align="left" valign="top"><!--<img src="image/Adv/indextop1.jpg" width="300" height="55" >-->
    <a href="login/login.asp">登录</a> <a href="login/Reg.asp">注册</a> <a href="product/check.asp">购物车</a> <a href="user/user_center.asp">我的账户</a></td>
    <td align="center"><img src="image/logo/logo.jpg" width="180" height="45"></td>
    <td width="200" align="right" valign="top"><a href="user/user_center.asp">订单查询</a> <a href="Contact/Contact.asp?id=2">付款说明</a> <a href="Contact/Contact.asp?id=9">退换细节</a></td>
  </tr>
  <tr>
    <td colspan="3" align="center" valign="top"><a target="blank" href="tencent://message/?uin=1157893749&amp;Site=代理咨询&amp;Menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=1:1157893749:4" alt="代理咨询" />代理咨询</a> <a target="blank" href="tencent://message/?uin=371528837&amp;Site=代理咨询&amp;Menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=1:371528837:4" alt="代理咨询" />代理咨询</a> <a target="blank" href="tencent://message/?uin=406478926&amp;Site=代理咨询&amp;Menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=1:406478926:4" alt="批发咨询" />批发咨询</a> <a target="blank" href="tencent://message/?uin=1123052852&amp;Site=代理咨询&amp;Menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=1:1123052852:4" alt="售后咨询" />售后咨询</a>
<a target="_blank" href="http://amos.im.alisoft.com/msg.aw?v=2&uid=myselfed_y&site=cntaobao&s=2&charset=utf-8" ><img border="0" src="http://amos.im.alisoft.com/online.aw?v=2&uid=myselfed_y&site=cntaobao&s=2&charset=utf-8" alt="有事点这里" />售后咨询</a>
	</td>
  </tr>
</table>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="6"><img src="image/top/top4.jpg" width="6" height="59" /></td>
    <td width="965" valign="middle" background="image/top/top5.jpg"><table width="940" height="58" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="30" colspan="2">
          <%
	SubLar=trim(request("LarCode"))
	set rs_MidCode=server.createobject("adodb.recordset") 
	sql_MidCode="select LarCode,MidCode from dk501_productclass where LarCode = '"& SubLar &"'order by MidSeq" 
	rs_MidCode.Open sql_MidCode,conn,1,1 	
	if rs_MidCode.bof and rs_MidCode.eof then
		response.write "<span class='White'>欢迎您来到DK501（裁难）！<a href=login/login.asp class='whit'>[登陆]</a> | <a href=login/Reg.asp class='whit'>[注册]</a> 本站全面支持支付宝和财付通付款！</span>"
	else
		Do While Not rs_MidCode.eof	
			response.write "&nbsp;<a class='whit' href=product.asp?LarCode="&Server.URLEncode(Trim(rs_MidCode("LarCode")))&"&MidCode="&Server.URLEncode(Trim(rs_MidCode("MidCode")))&">"&Trim(rs_MidCode("MidCode"))&"&nbsp;</a><span class='White'>|</span>"
		rs_MidCode.movenext
		  if rs_MidCode.eof then exit do
		loop		
	end if
	rs_MidCode.close
	set rs_MidCode=nothing
	%></td>
      </tr>
      <tr>
        <td width="702" height="28" valign="top"><table width="530" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <form action="product/product.asp" method="get" name="search" id="search">
              <td width="40" height="25" align="center"><%
	set rs_1=server.createobject("adodb.recordset") 
	sql="select distinct LarCode,LarSeq from dk501_productclass order by LarSeq" 
	rs_1.Open sql,conn,1,1 
		response.write "<select name='fenlei' class=5354_inputcolor>"
		response.write	"<option value='' selected>所有类别</option>"
	if rs_1.bof and rs_1.eof then
		response.write "<option selected value=''>暂无类别</option>"
	else
		Do While Not rs_1.eof	
		response.write "<option value='"&rs_1("larCode")&"'>"&rs_1("LarCode")&"</option>"
		rs_1.movenext
		  if rs_1.eof then exit do
		loop
		response.write "<option value='未归类'>其它</option>"
	end if
	rs_1.close
	set rs_1=nothing
	%></td>
              <td width="336" align="center"><input name="name_key" type="text" id="name_key"  style="width:280px;" value="<%=trim(request("name_key"))%>" maxlength="140" /></td>
              <td width="63" height="25"><input type="submit" name="button32" id="button32" value="搜索" class="search-input"/></td>
              <td width="91" align="center" valign="bottom" nowrap="nowrap">&nbsp;</td>
            </form>
            </tr>
        </table></td>
        <td width="258" valign="top"><table width="100" border="0" align="right" cellpadding="0" cellspacing="0">
          <tr>
            <td><a href="product/check.asp"><img src="image/top/top11.jpg" width="73" height="27" border="0" /></a></td>
            <td><a href="user/user_center.asp"><img src="image/top/top12.jpg" width="59" height="27" border="0" /></a></td>
            <td><a href="user/user_center.asp"><img src="image/top/top13.jpg" width="83" height="27" border="0" /></a></td>
            <td><a href="user/user_center.asp"><img src="image/top/top14.jpg" width="43" height="27" border="0" /></a></td>
          </tr>
        </table></td>
        </tr>
    </table></td>
    <td width="6" align="right"><img src="image/top/top6.jpg" width="6" height="59" /></td>
  </tr>
</table>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="209" height="158" align="left" valign="top"><table border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="41"><!--#include file="include/web_login.asp"--></td>
      </tr>
      
      <tr>
        <td height="93" valign="middle"><img src="image/adv/index_c_1.gif" width="195" height="85"></td>
      </tr>
    </table></td>
	<td>
	<a target=_self href="javascript:goUrl()" class="b-top b-left b-right b-bottom">
	<script type="text/javascript">imgUrl1="image/adv/g1.jpg";
imgtext1="宣传"
imgLink1="Contact/Contact.asp?id=2"
imgUrl2="image/adv/g2.jpg";
imgtext2="联系我们"
imgLink2="Contact/Contact.asp?id=2"
imgUrl3="image/adv/g3.jpg";
imgtext3="代理需知"
imgLink3="Contact/Contact.asp?id=2"
imgUrl4="image/adv/g4.jpg";
imgtext4="退换说明"
imgLink4="Contact/Contact.asp?id=2"
imgUrl5="image/adv/g5.jpg";
imgtext5="常见问题"
imgLink5="Contact/Contact.asp?id=2"

 var focus_width=520
 var focus_height=220
 var text_height=0
 var swf_height = focus_height+text_height
 
 var pics=imgUrl1+"|"+imgUrl2+"|"+imgUrl3+"|"+imgUrl4+"|"+imgUrl5
 var links=imgLink1+"|"+imgLink2+"|"+imgLink3+"|"+imgLink4+"|"+imgLink5
 var texts=imgtext1+"|"+imgtext2+"|"+imgtext3+"|"+imgtext4+"|"+imgtext5
 
 document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
 document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="image/focus.swf"><param name="quality" value="high"><param name="bgcolor" value="#F0F0F0">');
 document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
 document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
 document.write('<embed src="image/flash/focus.swf" wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" bgcolor="#F0F0F0" quality="high" width="'+ focus_width +'" height="'+ focus_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');  document.write('</object>');
 </script>
 </a></td>
    <td width="237" valign="top">
    <table width="100" border="0" align="right" cellpadding="0" cellspacing="0">
      <tr>
        <td><img src="image/adv/index1.jpg" width="232" height="63" /></td>
      </tr>
      <tr>
        <td height="96" valign="middle" bgcolor="#E3E3E5"><a href="<%=hf3url%>" title="<%=hf3tit%>" target="_blank"></a><a href="product/product.asp?LarCode=%C4%D0%D7%B0"><img src="image/adv/index_c_3.jpg" width="230" height="90" border="0"></a></td>
      </tr>
      <tr>
        <td valign="top" bgcolor="#E3E3E5"><a href="product/product.asp?LarCode=%C4%D0%D7%B0"><img src="image/adv/index_c_2.jpg" width="225" height="65" border="0"></a></td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="124" align="center" background="image/top/bg_1.jpg"><font color="#FFFFFF">最新上架</font></td>
    <td width="852" background="image/top/bg_2.jpg">&nbsp;20010年的春天,你打算穿什么呢？ 流行的、时尚的、前沿的····</td>
    <td width="3" align="center"><img src="image/top/bg_3.jpg" width="4" height="26" /></td>
  </tr>
</table>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0" >
        <tr>
          <%set rs=server.CreateObject("adodb.recordset")
		  rs.open "select top 12 ProdId,ImgPrev,ProdName,Prod1,PriceOrigin,PriceList,AddDate from dk501_product where online=2 order by AddDate desc ",conn,1,1
		  if rs.eof and rs.bof then
		  response.write "<center><br><font color=red size=2>对不起，暂无此类商品！</font></font>"
		  'response.End
		  else
		  %>
          <%
		if not rs.eof then
		i=1
		do while not rs.eof%>
          <td height="110" valign="top" style="padding-top:4px;" ><table border="0" cellspacing="0" cellpadding="0" valign="middle" width="140">
            <tr>
              <td width="116" height="116" align="center" valign="middle" background="images/bg/1.jpg" class="b-top b-left b-right b-bottom"><a href='product/list.asp?ProdId=<%=rs("ProdId")%>' target="_blank"><img src="<%=rs("ImgPrev")%>" alt="<%=rs("ProdName")%>" width="150" height="150"  border="0" /></a></td>
            </tr>
            <tr>
              <td height="16" align="left"><a href='product/list.asp?ProdId=<%=rs("ProdId")%>' target="_blank"><%=lleft(rs("ProdName"),14)%></a></td>
            </tr>
            <tr>
              <td height="16" align="left"><font color="#DB3578">尺码：<%=Left(Trim(rs("Prod1")),16)%></font></td>
            </tr>
            <tr>
              <td height="16" align="left">市场价：￥<%=formatnumber(rs("PriceOrigin"),2)%></td>
            </tr>
          </table></td>
          <%if i mod 6 = 0 then%>
        </tr>
        <%end if
	    rs.movenext
         i=i+1
    	 loop
		rs.close
		end if
		end if 
	%>
</table>
<!--#include file="include/login/bootom.asp"-->
</body>
</html>
<% call CloseConn() %>