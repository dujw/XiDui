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
	response.write "alert('�����ˣ�û����������ţ���������Ȩ�鿴�˶�����');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
else 
Recname2=rs("Recname2")
Recaddress2=rs("Recaddress2")
'��ݷ�ʽ
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
      <TD height="30" align="left" bgColor="#f0f0f0">�����ţ� <strong><font color=red><%=rs("OrderNum")%></font></strong> ���ڣ�<%=rs("ordertime")%> �û����� <%=request.cookies("buyok")("userid")%></TD>
    </TR>
	
<%
Set rs_user = conn.Execute("select UserName,HomePhone,Address from dk501_user where UserId='"&request.cookies("buyok")("userid")&"' ") 
if not (rs_user.eof and rs_user.bof) then
Do while not rs_user.eof
%>
    <TR>
      <TD bgColor="#f0f0f0">�� �� �ˣ�<%=rs_user("UserName")%></TD>
    </TR>
    <TR>
      <TD bgColor="#ffffff"><TABLE id="table1" cellSpacing="0" cellPadding="0" width="600" border="0">
          <TR>
            <TD width="238" height="23" align="left" bgColor="#ffffff">��ʵ������<%=rs_user("UserName")%></TD>
            <TD align="left" bgColor="#ffffff">�绰��<%=rs_user("HomePhone")%> </TD>
          </TR>
          <TR>
            <TD height="23" colspan="2" align="left" bgColor="#ffffff">��ϸ��ַ��<%=rs_user("Address")%></TD>
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
      <TD noWrap bgColor="#f0f0f0">������ʽ��<%=kuaidi%> ���׺ţ�<%=jiaoyihao%></TD>
    </TR>
    <TR>
      <TD bgColor="#ffffff"><TABLE id="table1" cellSpacing="0" cellPadding="0" width="600" border="0">
        
          <TR>
            <TD height="23" noWrap bgColor="#ffffff">������ʽ��<%=kuaidi%>  &nbsp;<%=memo%></TD>
          </TR>
        
      </TABLE></TD>
    </TR>
    <TR>
      <TD bgColor="#f0f0f0">�� �� �ˣ�<strong><%=rs("RecName2")%></strong></TD>
    </TR>
    <TR>
      <TD bgColor="#ffffff"><TABLE id="table1" cellSpacing="0" cellPadding="0" width="600" border="0">
        
          <TR>
            <TD noWrap width="50%" bgColor="#ffffff" height="25">������֣� <strong><%=rs("RecName2")%></strong>  �ƶ��绰�� <%=rs("RecPhone2")%></TD>
        </TR>
          <TR>
            <TD width="50%" bgColor="#ffffff" height="25">�ջ���ַ�� <%=rs("RecAddress2")%></TD>
        </TR>
          <TR>
            <TD bgColor="#EEEEEE" height="25">�˿�˵����<%=rs("Notes")%></TD>
          </TR>
        
      </TABLE></TD>
    </TR>
    <TR bgColor="#f0f0f0">
      <TD align="left" bgColor="#f0f0f0">���������б� </TD>
    </TR>
    <TR>
      <TD bgColor="#ffffff"><TABLE width="590" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
 <%
 Total = 0
 Sum_youfei = csng(rs("youfei"))
 do while not rs.eof 
 Sum = csng(rs("BuyPrice"))*csng(rs("ProdUnit")) 
 Total = Sum + Total '�����ܽ��
 %>
          <TR>
            <TD width="140" rowSpan="5"><a href="/product/list.asp?ProdId=<%=rs("Prodid")%>" target="_blank"><img src=<%=rs("ImgPrev")%> width="140" height="140" border="0"></a></TD>
            <TD width="457" bgcolor="#FFFFFF">&nbsp;���ƣ�<a href="/product/list.asp?ProdId=<%=rs("Prodid")%>" target="_blank"><%=rs("ProdName")%></a></TD>
          </TR>
          <TR>
            <TD bgcolor="#FFFFFF">&nbsp;���룺<%=rs("xinhao")%></TD>
          </TR>
          <TR>
            <TD bgcolor="#FFFFFF">&nbsp;������<%=rs("ProdUnit")%></TD>
          </TR>
          <TR>
            <TD bgcolor="#FFFFFF">&nbsp;���ۣ�<%=formatnum(csng(rs("BuyPrice")),2)%></TD>
          </TR>
          <TR>
            <TD bgcolor="#FFFFFF">&nbsp;�ϼƣ�<%=formatnum(sum,2)%></TD>
          </TR>
          
                               <%
rs.movenext
loop
end if
%>
      </TABLE></TD>
    </TR>
    <TR>
      <TD align="right" bgColor="#ffffff"><p>&nbsp;��Ʒ�ܼۣ�<%=formatnum(total,2)%>Ԫ �ʷѣ�<%=Sum_youfei%> ��Ʒ��������<%=zongzhong%></p></TD>
    </TR>  
</TABLE>
<% call CloseConn() %>