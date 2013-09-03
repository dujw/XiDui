<!--#include file="include/connect.asp" -->
<%
call aspsql()
'判断用户是否登陆
call checklogin()
Name   = request.cookies("buyok")("userid")

Pay_Mode1=request.form("Pay_Mode") '可以不要
		Recname2=Trim(request.form("Recname2"))
			If Recname2="" Then
				Response.Write("<script language=""JavaScript"">alert(""提醒：请输入收货人的姓名！"");history.go(-1);</script>")
				response.end
			end if
		Recaddress2=Trim(request.form("Recaddress2"))
			If Recaddress2="" Then
				Response.Write("<script language=""JavaScript"">alert(""提醒：请输入收货人的地址！"");history.go(-1);</script>")
				response.end
			end if
		RecPhone2=Trim(request.form("RecPhone2"))
			If RecPhone2="" Then
				Response.Write("<script language=""JavaScript"">alert(""提醒：请输入收货人的联系方式！"");history.go(-1);</script>")
				response.end
			end if
		Zipcode2=request.form("ZipCode2")
'==================开始支付方式的添加====================
		kuaidi=Trim(request.form("kuaidi"))            '快递方式
		If kuaidi="" Then
				Response.Write("<script language=""JavaScript"">alert(""提醒：请选择快递货运方式！"");history.go(-1);</script>")
				response.end
			end if
		
		youfei=Cint(Trim(request.form("youfei")))      '邮费费用
		If youfei="" Then
			Response.Write("<script language=""JavaScript"">alert(""提醒：请输入真确的邮费！"");history.go(-1);</script>")
			response.end
		end if
		alipay=request.form("alipay")               '支付方式		  
		Transactions=Trim(request.form("Transaction"))  '支付宝的帐号
		If Transactions="" Then
			Response.Write("<script language=""JavaScript"">alert(""提醒：请输入支付宝号！"");history.go(-1);</script>")
			response.end
		end if
		
		set rs1=server.createobject("adodb.recordset")
		sql1="select jiaoyihao from dk501_borderList where jiaoyihao ='"&Transactions&"'"
		rs1.open sql1,conn,1,3
		if not rs1.eof then
				Response.Write("<script language=""JavaScript"">alert(""提醒：交易号已经存在！"");history.go(-1);</script>")
				response.end
		else
			Transactions=Trim(request.form("Transaction"))
		end if
'========================================================
jifentotle=request.form("jifensum")
Notes=request.form("Notes")
post=request.form("post")       '可以不要
Gettime=request.form("gettime") '可以不要
if reg<>"1" then
	if request.cookies("buyok")("userid")="" then 
		response.redirect "../error/error.asp?error=006"
		response.end
	end if
end if
%>

<%
randomize
	d=right("00"&int(99*rnd()),2)
	yy=right(year(date),2)
	mm=right("00"&month(date),2)
	dd=right("00"&day(date),2)
	riqi=yy & mm & dd
	xiaoshi=right("00"&hour(time),2)
	fenzhong=right("00"&minute(time),2)
	miao=right("00"&second(time),2)
	inBillNo="XD-"&yy & mm & dd & xiaoshi & fenzhong & miao & d
	inBillNo1=yy&"年" & mm &"月"& dd &"日"& xiaoshi& "时"& fenzhong &"分"
%>

<%
'===================================产品原来后台设置的价格====================


guahao=0
fei=0
shouzhong=0
xuzhong=1
xzf=1
'===================================================================================

Name=request.cookies("buyok")("userid")
	RecName2=request.form("RecName2")
	RecAddress2=request.form("RecAddress2")
	Zipcode2=request.form("ZipCode2")
	RecPhone2=request.form("RecPhone2")
	Notes=request.form("Notes")
	'buylist=request.form("buylist")
	buylist = request.cookies("buyok")("cart")
	userkou=checkuserkou()
		if reg<>"1" and request.cookies("buyok")("userid")="" then
			response.redirect "../error/errorerror.asp?error=006"
			response.end
		end if
conn.execute("update dk501_border set OrderNum = '"&inBillNo&"'  where UserId='"& Name &"' and OrderNum='"& Name &"'")
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
        <td height="65" align="left" background="../image/order/news2.jpg" class="b-left b-right">&nbsp;&nbsp;<img src="image/order/user1.gif" width="11" height="11" /> 当前位置：<a href="index.asp">首页</a> 》购物中心 》我的订单</td>
      </tr>
      <tr>
        <td height="28" align="left" valign="middle" background="../image/order/bucou1.jpg" class="b-left b-right">&nbsp;&nbsp;&nbsp;<font color="#FFFFFF">订单提交步筹：1.选择产品订购产品 2.填写要的尺码（继续购买其它产品）3.填写联系方式以及付款 4.成功提交订单</font></td>
  </tr>
      <tr>
        <td height="89" align="left" valign="top" class="b-left b-right b-bottom" style="padding-top:10px; padding-right:10px; padding-bottom:10px; padding-left:10px;"><table width="800" height="50" align="center" cellpadding="0" cellspacing="0" class="jiesuan">
                <tr>
                  <td width="565"><table width="800"  border="0" cellpadding="0" cellspacing="1" bgcolor="#D1D2D3">
                      <tr bgcolor="#FFFFFF">
                        <td width="126" height="25" align="right">您的定单编号为：</td>
                        <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="6%">&nbsp;</td>
                              <td width="94%" height="25"><span class="bold"><%=inBillNo%></span>（查询订单请记住定单号）</td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td width="126" height="25" align="right">您提交定单时间：</td>
                        <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="6%">&nbsp;</td>
                              <td width="94%" height="24"><%=inBillNo1%></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td height="28" align="right" bgcolor="#FFFFFF">收货人的姓名：&nbsp;</td>
                        <td width="27" bgcolor="#FFFFFF">&nbsp;</td>
                        <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D1D2D3">
                            <tr>
                              <td width="23%" height="25" align="right" bgcolor="#FFFFFF">收货人的姓名：</td>
                              <td height="25" bgcolor="#FFFFFF"><%=Trim(request.form("Recname2"))%></td>
                            </tr>
                            <tr>
                              <td height="25" align="right" bgcolor="#FFFFFF">收货人的电话：</td>
                              <td height="25" bgcolor="#FFFFFF"><%=Trim(request.form("RecPhone2"))%></td>
                            </tr>
                            <tr>
                              <td height="25" align="right" bgcolor="#FFFFFF">收货人的地址：</td>
                              <td height="25" bgcolor="#FFFFFF"><%=Trim(request.form("Recaddress2"))%></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td height="28" rowspan="3" align="right">您的订单金额：&nbsp;</td>
                        <td width="27" rowspan="3">&nbsp;</td>
                        <td width="138" height="25" align="right">商品总计金额：</td>
                        <td width="472" height="25">￥<font color="#FF0000"><%=Trim(request.form("sumkiss"))%></font>元</td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td height="25" align="right">发货邮费：</td>
                        <td width="472" height="25">￥<%=Trim(request.form("youfei"))%>元</td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td height="22" align="right">总付金额：</td>
                        <td height="22">￥<%=formatnum(Trim(request.form("sumkiss"))+Cint(Trim(request.form("youfei"))),2)%></td>
                      </tr>
                  </table></td>
                </tr>
            </table></td>
      </tr>
</table>	  </td>
    </tr>
  </table>
  <%
set rsadd=server.createobject("adodb.recordset")
sql="select * from dk501_borderList "
rsadd.open sql,conn,1,3
	rsadd.AddNew 
		rsadd("UserId")=Name
		rsadd("OrderNum")=inBillNo
		rsadd("OrderTime")=Now()
		rsadd("LastModifytime")=Now()
		rsadd("OrderSum")=FormatNum(Trim(request.form("sumkiss")),2)
		rsadd("RecName2")=RecName2
		rsadd("RecAddress2")=RecAddress2
		rsadd("RecPhone2")=RecPhone2
		rsadd("ZipCode2")=ZipCode2
		rsadd("zhongliang")=Trim(request.form("zongzhong"))
		'==================================
		rsadd("kuaidi")=kuaidi
		rsadd("youfei")=youfei
		rsadd("alipay")=alipay
		rsadd("jiaoyihao")=Transactions
		rsadd("Notes")=Notes
		rsadd("Del")=0
		'==================================
	rsadd.Update
rsadd.close
set rsadd=nothing
%>
<%
response.cookies("buyok")("cart")=""
response.cookies("buyok")("total")=""
%>
</div>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>
