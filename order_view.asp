<!--#include file="include/connect.asp" -->
<%
call aspsql()
call checklogin()
OrderNum=request("OrderNum")

dim jia
sqlinfo = "select A.OrderNum,A.ProdId,A.BuyPrice,A.xinhao,A.ProdUnit,B.kuaidi,B.jiaoyihao,B.youfei,B.Notes,B.ordertime,B.RecName2,B.memo,B.RecPhone2,B.RecAddress2,B.zhongliang,C.ImgPrev,C.ProdName "&_
  " from dk501_border A,dk501_borderList B,dk501_product C"&_
  " where b.del<>1 and A.UserId='"&request.cookies("buyok")("userid")&"' and A.OrderNum='"&OrderNum&"' and A.OrderNum = B.OrderNum and A.ProdId=C.Prodid"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sqlinfo,conn,1,1
if rs.eof and rs.bof  then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，没有这个订单号，或者您无权查看此订单！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
else 
Recname2=rs("Recname2")
Recaddress2=rs("Recaddress2")
'快递方式
kuaidi = rs("kuaidi")
jiaoyihao = rs("jiaoyihao")
memo = rs("memo")
RecPhone2=rs("RecPhone2")
inBillNo=rs("OrderNum")
zongzhong=rs("zhongliang")
%>



<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
<TABLE width="600" border="0" align="center" cellPadding="5" cellSpacing="1" bgColor="#c8c8c8">
  
    <TR>
      <TD height="30" align="left" bgColor="#f0f0f0">定单号： <strong><font color=red><%=rs("OrderNum")%></font></strong> 日期：<%=rs("ordertime")%> 用户名： <%=request.cookies("buyok")("userid")%></TD>
    </TR>
	
<%
Set rs_user = conn.Execute("select UserName,HomePhone,Address from dk501_user where UserId='"&request.cookies("buyok")("userid")&"' ") 
if not (rs_user.eof and rs_user.bof) then
Do while not rs_user.eof
%>
    <TR>
      <TD bgColor="#f0f0f0">订 货 人：<%=rs_user("UserName")%></TD>
    </TR>
    <TR>
      <TD bgColor="#ffffff"><TABLE id="table1" cellSpacing="0" cellPadding="0" width="600" border="0">
          <TR>
            <TD width="238" height="23" align="left" bgColor="#ffffff">真实姓名：<%=rs_user("UserName")%></TD>
            <TD align="left" bgColor="#ffffff">电话：<%=rs_user("HomePhone")%> </TD>
          </TR>
          <TR>
            <TD height="23" colspan="2" align="left" bgColor="#ffffff">详细地址：<%=rs_user("Address")%></TD>
          </TR>
        
      </TABLE></TD>
    </TR>
	<%
	rs_user.movenext
	loop
	end if  
	rs_user.close
	set rs_user=nothing
	%>
    <TR>
      <TD noWrap bgColor="#f0f0f0">发货方式：<%=kuaidi%> 交易号：<%=jiaoyihao%></TD>
    </TR>
    <TR>
      <TD bgColor="#ffffff"><TABLE id="table1" cellSpacing="0" cellPadding="0" width="600" border="0">
        
          <TR>
            <TD height="23" noWrap bgColor="#ffffff">发货方式：<%=kuaidi%>  &nbsp;<%=memo%></TD>
          </TR>
        
      </TABLE></TD>
    </TR>
    <TR>
      <TD bgColor="#f0f0f0">收 货 人：<strong><%=rs("RecName2")%></strong></TD>
    </TR>
    <TR>
      <TD bgColor="#ffffff"><TABLE id="table1" cellSpacing="0" cellPadding="0" width="600" border="0">
        
          <TR>
            <TD noWrap width="50%" bgColor="#ffffff" height="25">买家名字： <strong><%=rs("RecName2")%></strong>  移动电话： <%=rs("RecPhone2")%></TD>
        </TR>
          <TR>
            <TD width="50%" bgColor="#ffffff" height="25">收货地址： <%=rs("RecAddress2")%></TD>
        </TR>
          <TR>
            <TD bgColor="#EEEEEE" height="25">顾客说明：<%=rs("Notes")%></TD>
          </TR>
        
      </TABLE></TD>
    </TR>
    <TR bgColor="#f0f0f0">
      <TD align="left" bgColor="#f0f0f0">订购货物列表 </TD>
    </TR>
    <TR>
      <TD bgColor="#ffffff"><TABLE width="590" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
 <%
 Total = 0
 Sum_youfei = csng(rs("youfei"))
 do while not rs.eof 
 Sum = csng(rs("BuyPrice"))*csng(rs("ProdUnit")) 
 Total = Sum + Total '计算总金额
 %>
          <TR>
            <TD width="140" rowSpan="5"><a href="/product/list.asp?ProdId=<%=rs("Prodid")%>" target="_blank"><img src=<%=rs("ImgPrev")%> width="140" height="140" border="0"></a></TD>
            <TD width="457" bgcolor="#FFFFFF">&nbsp;名称：<a href="/product/list.asp?ProdId=<%=rs("Prodid")%>" target="_blank"><%=rs("ProdName")%></a></TD>
          </TR>
          <TR>
            <TD bgcolor="#FFFFFF">&nbsp;尺码：<%=rs("xinhao")%></TD>
          </TR>
          <TR>
            <TD bgcolor="#FFFFFF">&nbsp;数量：<%=rs("ProdUnit")%></TD>
          </TR>
          <TR>
            <TD bgcolor="#FFFFFF">&nbsp;单价：<%=formatnum(csng(rs("BuyPrice")),2)%></TD>
          </TR>
          <TR>
            <TD bgcolor="#FFFFFF">&nbsp;合计：<%=formatnum(sum,2)%></TD>
          </TR>
          
                               <%
rs.movenext
loop
end if
%>
      </TABLE></TD>
    </TR>
    <TR>
      <TD align="right" bgColor="#ffffff"><p>&nbsp;商品总价：<%=formatnum(total,2)%>元 邮费：<%=Sum_youfei%> 商品总重量：<%=zongzhong%></p></TD>
    </TR>  
</TABLE>
<% call CloseConn() %>