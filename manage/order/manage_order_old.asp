<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<!--#include file="../../include/md5.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
<script language=javascript>
<!--
function CheckAll(form){
for (var i=0;i<form.elements.length;i++){
var e = form.elements[i];
if (e.name != 'chkall') e.checked = form.chkall.checked; 
}
}
-->
</script>
</head>
<BODY>
<%

action = request("action")
if action="" then call order_index()
if action="list" then call Order()
if action="删除" then call orderdel()
sub order_index()

%>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD bgColor=#f4f5fd height=30><font color="#000000">　<B><font color="#FF0000">[定单管理]</font></B></font>&nbsp;&nbsp;&nbsp;
      <%				 
	sqlBigClass="select Status,StatusDefines from dk501_types where cateids=1 "
	Set rsBigClass= Server.CreateObject("ADODB.Recordset")
	rsBigClass.open sqlBigClass,conn,1,1
	if rsBigClass.bof and rsBigClass.eof then 
		response.Write "&nbsp;没有任何栏目"
	else	
		do while not rsBigClass.eof
	%>
      <a href="manage_order_old.asp?Status=<%=rsBigClass("Status")%>&show=yes"><%=rsBigClass("StatusDefines")%></a>&nbsp;|&nbsp;
      <%
		rsBigClass.movenext
		loop
	end if 
	rsBigClass.close
	Set rsBigClass=Nothing
	%>
    </TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table width="99%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
  <form action=manage_order1.asp method=post name=orderlist>
    <%
ordersql=request.cookies("buyok")("ordersql")
Recname2=Trim(request("Recname2"))
Status=Trim(request("Status"))
if request("show")="yes" then
  response.cookies("buyok")("ordersql")=""
end if
pages = 20
set rs = server.createobject("adodb.recordset")
'sql = "select a.ID,a.OrderNum,a.del,a.OrderSum,a.Recname2,a.kuaidi,a.youfei,a.alipay,a.jiaoyihao,a.RecPhone2,a.RecAddress2,a.ZipCode2,a.UserID,a.OrderTime,b.Status,b.StatusDefines from dk501_borderList a,dk501_types b where a.del=0 and a.Status=b.Status "
sql="select * from DU_Borderlist_v_2012 where del=0 "
	'开始判断订单状态	
	if Status <> "" then 
		sql = sql & " and Status = '" & Status & "'"
	end if
	'开始判断买家的名字
	if Recname2 <> "" then 
sql = sql & " and (Recname2 like '%" & Recname2 & "%' or userid like '%" & Recname2 & "%' or OrderNum like '%" & Recname2 & "%' or jiaoyihao like '%" & Recname2 & "%')"
	end if
if request("show")<>"yes" then 
	sql=sql+ordersql 
end if
	sql= sql+" order by OrderTime desc"
rs.open sql,conn,1,3
response.write "<!--"+SQL+"-->"
if rs.eof and rs.bof then
	response.write "<tr bgcolor=#FFFFFF><td colspan=7 align=center height=50>暂时没有订单</td></tr>"
response.end 
rs.close
set rs=nothing
end if
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
%>
    <tr bgcolor="#f4f5fd">
      <td width=5% height="20" align=center>选</td>
      <td height="20" align=center>订单号</td>
      <td height="20" align=center>金额</td>
      <td height="20" align=center>会员ID</td>
      <td align=center nowrap>买家信息</td>
      <td height="20" colspan="2" align=center nowrap>快递方式</td>
      <td height="20" colspan="2" align=center>付款方式</td>
      <td height="20" align=center>下单时间</td>
      <td height="20" align=center>订单状态</td>
    </tr>
    <%
Total = 0
Do While Not rs.eof and pages>0
	
	Sum = formatnum(csng(rs("OrderSum"))+csng(rs("youfei")),2)
	Sum = FormatNumber(Sum,2)
	Total = Sum + Total
%>
    <tr bgcolor=#FFFFFF onMouseOver="this.style.backgroundColor='#FFF7D2'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
      <td height="22"><input type='checkbox' value=<%=rs("ID")%> name=id></td>
      <td height="22" align=center><a alt='查看/编辑订单' href=manage_order1.asp?action=list&ID=<%=rs("OrderNum")%>><%=rs("OrderNum")%></a></td>
      <td height="22" align=center><%=formatnum(csng(rs("OrderSum"))+csng(rs("youfei")),2)%>&nbsp;</td>
      <td height="22" align=center><a alt='查看会员资料' href=../user/manage_user_edit.asp?action=useredit&id=<%=rs("userid")%>><%=rs("userid")%></a></td>
      <td height="22" align=center><%=rs("Recname2")%></td>
      <td height="22" align=center><%=rs("kuaidi")%></td>
      <td height="22" align=center><%=rs("youfei")%></td>
      <td height="22" align=center nowrap><%=Trim(rs("alipay"))%></td>
      <td height="22" align=center><%=rs("jiaoyihao")%></td>
      <td height="22" align=center nowrap><%=rs("OrderTime")%></td>
      <td height="22" align=center nowrap><%
if rs("Status")="0" then
	response.write "<font color=red>"&rs("StatusDefines")&"</font>"
elseif rs("Status")="99" then 
	response.write rs("StatusDefines")
elseif rs("Status")="88" then 
	response.write "<font color=#669900>"&rs("StatusDefines")&"</font>"
else
	response.write "<font color=#3300FF>"&rs("StatusDefines")&"</font>"
end if
%></td>
    </tr>
    <%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
%>
    <tr bgcolor="#FFFFFF">
      <td colspan=11><input type='checkbox' name=chkall onclick='CheckAll(this.form)'>
        <input type="submit" name="action" value="删除" onClick="{if(confirm('提示：您确实删除所选定的订单吗？')){this.document.orderlist.submit();return true;}return false;}">
      </td>
    </tr>
  </form>
</table>
<table width="99%" border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F" class="border">
  <tr class="tdbg">
    <form name="searchsoft" method="get" action="manage_order_old.asp">
      <td height="30" bgcolor="#F4F5FD"><strong>&nbsp;查找数据：</strong>
        <input name="Recname2" type="text" class=smallInput id="Recname2" size="30" maxlength="50">
        <input name="Query" type="submit" id="Query" value="查 询">
        <%=formatnumber(total,2)%></td>
    </form>
  </tr>
</table>
<%
rs.close
set rs=nothing
'if allpages<=1 then exit sub
	response.write "总订单数<font color=red>"&totalrs&"</font> &nbsp; 每页<font color=red>20</font>&nbsp; "
if page = 1 then
	response.write "<font color=darkgray>首页 前页</font>"
else
	response.write "<a href=manage_order_old.asp?Recname2="&Recname2&"&Status="&Status&"&page=1>首页</a> <a href=manage_order_old.asp?Recname2="&Recname2&"&Status="&Status&"&page="&page-1&">前页</a>"
end if


action="manage_order_old.asp?Recname2="&Recname2&"&Status="&Status&""				
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
	response.write "<font color=darkgray> 下页 末页</font>"
else
	response.write " <a href=manage_order_old.asp?Recname2="&Recname2&"&Status="&Status&"&page="&page+1&">下页</a> <a href=manage_order_old.asp?Recname2="&Recname2&"&Status="&Status&"&page="&allpages&">末页</a>"
end if
	response.write " &nbsp;&nbsp;&nbsp;第"&page&"页 共"&allpages&"页"
end sub

sub Order()
OrderNum=request("ID")
if request.form("edit")="ok" then
	'==========================================开始编辑订单人得信息===================================
	set rs=Server.Createobject("ADODB.RecordSet")
	sql="select * from dk501_borderList where OrderNum='"&OrderNum&"'"
	rs.Open sql,conn,1,3
	rs("LastModifytime") = now()
	if trim(request("Memo"))     <>"" then rs("Memo")       = trim(request("Memo"))
	if trim(request("Memo1"))     <>"" then rs("Memo1")       = trim(request("Memo1"))
	if trim(request("Status"))     <>"" then rs("Status")       = trim(request("Status"))
	if trim(request("alipay"))     <>"" then rs("alipay")       = trim(request("alipay")) '支付宝帐号
	if trim(request("jiaoyihao"))     <>"" then rs("jiaoyihao")       = trim(request("jiaoyihao")) '支付宝号码
	if trim(request("kuaidi"))     <>"" then rs("kuaidi")       = trim(request("kuaidi")) '快地
	if trim(request("youfei"))     <>"" then rs("youfei")       = trim(request("youfei")) '支付宝号码
	if trim(request("RecPhone2"))     <>"" then rs("RecPhone2")       = trim(request("RecPhone2")) '买家电话
	if trim(request("RecAddress2"))     <>"" then rs("RecAddress2")       = trim(request("RecAddress2")) '买家地址
	if trim(request("Recname2"))     <>"" then rs("Recname2")       = trim(request("Recname2"))
	rs("OrderSum") = trim(Int(request("Price_Total")))

	
	'开始更新货得准确得发货时间
	if request("Status")="99" then rs("Success_time") = date()
	if request("Status")="99" then rs("pei") = request.cookies("buyok")("admin")
	'==========================================开始编辑订购产品信息===================================
	if request("Status") <> "99" and request("Status") <> "" then
		set rs_Product =Server.Createobject("ADODB.RecordSet")
		sql_Product="select * from dk501_border where OrderNum='"&OrderNum&"'"
		rs_Product.Open sql_Product,conn,1,3
		if not rs_Product.eof then
			do while not rs_Product.eof	
				'更改数量 ProdUnit
				conn.execute("update dk501_border set ProdUnit= '"&Cint(Trim(request("spec_"&rs_Product("ProdId")&"_"&rs_Product("xinhao")))) &"' where ProdId='"& rs_Product("ProdId") &"' and OrderNum='"&OrderNum&"' and xinhao = '"& Trim(request("cart_"&rs_Product("ProdId")&"_"&rs_Product("xinhao"))) &"'")	
				
				'更改购买产品价格 BuyPrice
				conn.execute("update dk501_border set BuyPrice= '"&Cint(Trim(request("BuyPrice_"&rs_Product("ProdId")&"_"&rs_Product("xinhao")))) &"' where ProdId='"& rs_Product("ProdId") &"' and OrderNum='"&OrderNum&"' and xinhao = '"& Trim(request("cart_"&rs_Product("ProdId")&"_"&rs_Product("xinhao"))) &"'")
				'更改购买产品型号 xinhao
				rs_Product.movenext
			loop
		end if		
		rs_Product.close
		set rs_Product=nothing
	end if 
	

	'==========================================开始编辑代理业绩信息===================================
	set rstjr=conn.execute("select * from dk501_user where UserId='"&request("userid")&"'")
	set rsjifensum=conn.execute("select * from dk501_border where UserId='"&request("userid")&"'")

	if request("Status")="99" then	'当订单处于完成得状态得时候！开始对会员加积分和销售件数和单数进行数据整合
	'response.Write request("Price_Total") 总的价格 没加邮费的
	'response.Write request("Counts_Total") 总的数量		
	conn.execute("update dk501_user set totalsum= totalsum+"&request("Price_Total")&" where userid='"&request("userid")&"'")	
	conn.execute("update dk501_user set jifensum= jifensum+"&request("jifensum")&" where userid='"&request("userid")&"'")	
	conn.execute("update dk501_user set ALL_Total= ALL_Total+"&request("Counts_Total")&" where userid='"&request("userid")&"'")
	conn.execute("update dk501_user set Month_danshu= Month_danshu+1 where userid='"&request("userid")&"'")
	conn.execute("update dk501_user set Current_Total= Current_Total+"&request("Counts_Total")&" where userid='"&request("userid")&"'")
	end if

	rs.update
	rs.close
	set rs=nothing

	response.write "<script language='javascript'>"
	response.write "alert('操作成功，您已经修改一个订单。');"
	response.write "location.href='manage_order_old.asp?action=list&id="&request("ID")&"';"
	response.write "</script>"
	response.end
end if


set rs=conn.execute("select A.ordertime,b.ordersum,A.Userid,A.OrderNum,A.jifensum,A.ProdId,A.BuyPrice,A.ProdUnit,A.xinhao,A.xzf,B.Memo,B.Memo1,B.alipay,B.Status,B.jiaoyihao,B.youfei,B.kuaidi,B.zhongliang,B.Recname2,b.PEI,B.RecPhone2,B.RecAddress2,B.Zipcode2,B.Notes,B.Success_time,C.ProdName,C.Product_Price,C.ImgPrev from dk501_border A,dk501_borderList B,dk501_product C where A.OrderNum='"&OrderNum&"' and A.OrderNum = B.OrderNum and A.ProdId=C.Prodid")

if rs.eof and rs.bof  then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，没有此订单。');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
%>
<%
xzf=rs("xzf")
zongzhong=rs("zhongliang")
%>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD bgColor=#f4f5fd height=30><font color="#000000"><B><font color="#FF0000">查看/编辑订单资料</font></B></font>
      <%				 
	sqlBigClass="select Status,StatusDefines from dk501_types where cateids=1 "
	Set rsBigClass= Server.CreateObject("ADODB.Recordset")
	rsBigClass.open sqlBigClass,conn,1,1
	if rsBigClass.bof and rsBigClass.eof then 
		response.Write "&nbsp;没有任何栏目"
	else	
		do while not rsBigClass.eof
	%>
      <a href="manage_order_old.asp?Status=<%=rsBigClass("Status")%>&show=yes"><%=rsBigClass("StatusDefines")%></a>&nbsp;|&nbsp;
      <%
		rsBigClass.movenext
		loop
	end if 
	rsBigClass.close
	set rsBigClass=Nothing		
	%></TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table width="99%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
  <form action=manage_order1.asp?action=list method=post name=orderedit>
    <INPUT TYPE="hidden" name=edit value=ok>
    <INPUT TYPE="hidden" name=Id value=<%=rs("OrderNum")%>>
    <tr bgcolor="#FFFFFF">
      <td width="18%" height="20" align='right'>订 单 号：</td>
      <td height="20"><input type="text" name="" value="<%=rs("OrderNum")%>"  readonly></td>
      <td height="20" align="right" class="800_left_red">提交时间：</td>
      <td height="20"><input type="text" name="Input" value="<%=rs("Ordertime")%>"  readonly></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height="20" align='right' bgcolor="#FFFFFF" class="800_pssm">订单金额：</td>
      <td width="24%" height="20"><input type="text" name="ordersum" value="<%=formatnumber(Cint(rs("Ordersum"))+Cint(rs("youfei")),2)%>"  readonly></td>
      <td width="14%" height="20" align="right" bgcolor="#FFFFFF" class="800_pssm">获得积分：</td>
      <td width="44%" height="20"><input type="text" name="jifensum" value="<%=rs("jifensum")%>"  readonly></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height="20" align='right' bgcolor="#FFFFFF" >付款方式</td>
      <td height="20" nowrap><input name="alipay" type="text" id="alipay" value="<%=rs("alipay")%>" size="8" maxlength="8">
        <input name="jiaoyihao" type="text" id="jiaoyihao" value="<%=rs("jiaoyihao")%>" size="28" maxlength="28"></td>
      <td width="14%" height="20" align="right" bgcolor="#FFFFFF" class="800_pssm">发货方式：</td>
      <td width="44%" height="20"><input name="kuaidi" type="text" id="kuaidi" value="<%=RS("kuaidi")%>" size="6" maxlength="6">
        <input name="youfei" type="text" id="youfei" value="<%=rs("youfei")%>" size="3" maxlength="3" >
        元
        <%youfei=rs("youfei")%>
        (申通  圆通  顺风  EMS) </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height="20" align='right'>会员ID：</td>
      <td height="20"><a href="1.asp?a=<%=FormatNumber(Total,2)%>&jiefen=<%=rs("jifensum")%>"></a>
        <input type="text" name="userid" value="<%=rs("userid")%>"  readonly>
        <a href="manage_user1.asp?action=useredit&id=<%=rs("userid")%>">查看用户信息</a></td>
      <td height="20" align="right">发货物人：</td>
      <td height="20"><input alt=顾客选择的配送方式 type="text" value="<%=rs("pei")%>" readonly >
        <%=rs("Success_time")%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align='right' bgcolor="#D9F7FF">&nbsp;</td>
      <td height="25" colspan="3" bgcolor="#D9F7FF" class="800_title">卖家（收货人）详细联系方式</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height="20" align='right' bgcolor="#EBEBEB"><font color="#FF0000">定货人送给：</font></td>
      <td height="20" bgcolor="#EBEBEB"><input type="text" name="Recname2" value="<%=rs("Recname2")%>" readonly maxlength="30"></td>
      <td height="20" align="right" bgcolor="#EBEBEB"><font color="#FF0000">联系电话：</font></td>
      <td height="20" bgcolor="#EBEBEB"><input name="RecPhone2" type="text" value="<%=Trim(rs("RecPhone2"))%>" size="26"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height="20" align='right' bgcolor="#EBEBEB"><font color="#FF0000">地址：</font></td>
      <td height="20" colspan="3" bgcolor="#EBEBEB"><input name="RecAddress2" type="text" value="<%=Trim(rs("RecAddress2"))%>" size="80"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height="20" align='right' bgcolor="#EBEBEB"><font color="#FF0000">邮政编码：</font></td>
      <td height="20" bgcolor="#EBEBEB"><input type="text" name="Zipcode2" value="<%=rs("Zipcode2")%>" readonly  maxlength="10"></td>
      <td height="20" colspan="2" align="right" bgcolor="#EBEBEB">&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align='right'>顾客说明：</td>
      <td colspan="3"><textarea name='Notes' cols="90" rows='2' readonly alt=顾客对此订单的其它要求或说明 ><%=rs("Notes")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align='right'>发货单号：</td>
      <td colspan="3" bgcolor="#FFFFFF"><textarea name='memo' cols="90" rows='3' style="overflow:auto; background-color: #FFFFCC;" alt="填写对订单处理的说明"><%=rs("Memo")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align='right'>订单意见：</td>
      <td colspan="3" bgcolor="#FFFFFF"><textarea name="memo1" cols="90" rows='3' id="memo1" style="overflow:auto; background-color: #EBEBEB;" alt="填写对订单处理的说明"><%=rs("Memo1")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=30 align='right'>订单状态：</td>
      <td colspan="3"><%
zhongliang=rs("zhongliang")
if rs("Status")="11" or  rs("Status")="12" or  rs("Status")="99" then
set rs2 = conn.Execute("select * from dk501_types where Status='"&rs("Status")&"' and cateids=1")
	response.write "<b>"&rs2("StatusDefines")&"<B>"
else
set rs2 = conn.execute("select * from dk501_types where cateids=1 Order by Status asc")
	response.write	"<select name='Status'>"
	Do While Not rs2.eof	
	if rs2("Status") = rs("Status") then 
	response.write "<option value='"&rs2("Status")&"' selected>"&rs2("StatusDefines")&"</option>"
	else
	response.write "<option value='"&rs2("Status")&"'>"&rs2("StatusDefines")&"</option>"
	end if
	rs2.movenext
	if rs2.eof then	exit do
	loop
	response.write "</select>"
set rs2=nothing
end if
%>
      </td>
    </tr>
    <tr bgcolor=#FFFFFF>
      <td align=right>购物清单：</td>
      <td colspan="3"><table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#b1bfee">
          <tr bgcolor="#FFFFFF">
            <td height="22" colspan="2" align="center" bgcolor="#FFFFFF">商品编号</td>
            <td width="41%" height="22" bgcolor="#FFFFFF">商品名称</td>
            <td height="22" nowrap bgcolor="#FFFFFF">购买数量</td>
            <td width="19%" align="center" nowrap>购买产品型号</td>
            <td width="8%" height="22" nowrap bgcolor="#FFFFFF">结算单价</td>
            <td width="5%" height="22" align=center nowrap bgcolor="#FFFFFF">总金额</td>
            <% if request.cookies("buyok")("admin") = "manage" or request.cookies("buyok")("admin") = "马勇" then %>
            <td width="7%" align=center bgcolor="#FFFFFF">默认</td>
            <% end if %>
          </tr>
          <%
	Total = 0
 do while not rs.eof 
 Sum = FormatNumber(csng(Cint(rs("BuyPrice")))*csng(Cint(rs("ProdUnit"))),2)
 Total = Sum + Total '计算总金额
 sum_ProdUnit=sum_ProdUnit+Cint(rs("ProdUnit"))
%>
          <tr bgcolor=#FFFFFF  onMouseOver="this.style.backgroundColor='#FFF732'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
            <td width="7%" height="22" align="center" bgcolor="#FFFFFF"><a href=../product/manage_product3.asp?action=detail&id=<%=rs("ProdId")%>><img src="../../<%=rs("ImgPrev")%>" width="50" height="50" border="0"></a></td>
            <td width="5%" align="center"><%=rs("ProdId")%>
              <INPUT TYPE="hidden" name=<%=rs("ProdId")%> value=<%=rs("ProdId")%>></td>
            <td height="22"><a alt='查看此商品的前台页面' href=../gotourl.asp?/product/list.asp?ProdId=<%=rs("ProdId")%> target='blank_'><%=rs("ProdName")%></a></td>
            <td height="22"><input name="spec_<%=rs("ProdId")%>_<%=rs("xinhao")%>" type="text"  value="<%=rs("ProdUnit")%>" size="4" maxlength="4">
              <input name="old_spec_<%=rs("ProdId")%>_<%=rs("xinhao")%>" type="hidden" value="<%=rs("ProdUnit")%>" size="4" maxlength="4">
            </td>
            <td height="22" align="center" ><input name="cart_<%=rs("ProdId")%>_<%=rs("xinhao")%>" type="text" value="<%=Trim(rs("xinhao"))%>" size="20"></td>
            <td height="22" align="center" ><input name="BuyPrice_<%=rs("ProdId")%>_<%=Trim(rs("xinhao"))%>" type="text" value="<%=FormatNumber(csng(rs("BuyPrice")),2)%>" size="4" maxlength="4"></td>
            <td height="22" align="center"><%=Sum%></td>
            <% if request.cookies("buyok")("admin") = "manage" or request.cookies("buyok")("admin") = "马勇" then %>
            <td height="22" align="center"><%=rs("Product_Price")%></td>
            <% end if %>
          </tr>
          <%
	rs.movenext
	loop
%>
          <TR bgcolor="#FFFFFF">
            <TD colspan='8'> 商品总价：<%=FormatNumber(Total,2)%>
              <input name="Price_Total" type="hidden" value="<%=Total%>">
              商品总数：<%=sum_ProdUnit%>
              <input name="Counts_Total" type="hidden" value="<%=sum_ProdUnit%>">
              商品总重：<%=zhongliang%> 总计费用：<font color=red><b><%=FormatNumber((Total+youfei),2)%></font></b> 元 </td>
          </TR>
        </table></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td colspan=4><input type='submit' name='action' value='修改订单'>
        <a href="1.asp?a=<%=FormatNumber(Total,2)%>&sum_ProdUnit=<%=sum_ProdUnit%>"></a> </td>
    </tr>
  </form>
</table>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<%
set rs=nothing
end sub

sub orderdel()
delid=request("id")
if Delid="" or isnull(Delid) then
response.write "<script language='javascript'>"
response.write "alert('操作失败，您什么也没有选择。');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from dk501_borderList where id in ("&Delid&")"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，该订单不存在或者已经被删除。');"		
	response.write "</script>"			
	response.end 
else
	Do while not rs.eof
	rs("del")=true
	rs.update
	rs.movenext
	loop
rs.close
set rs=nothing
	response.write "<script language='javascript'>"
	response.write "alert('操作成功，所选订单已被放入订单回收站，您可前往回收站彻底清除或恢复。');"
	response.write "location.href='manage_order_old.asp';"			
	response.write "</script>"
end if
end sub
%>
