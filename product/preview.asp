<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../include/connect.asp"-->
<%
call aspsql()
'�ж��û��Ƿ��½
call checklogin()
Name   = request.cookies("buyok")("userid")

Pay_Mode1=request.form("Pay_Mode") '���Բ�Ҫ
		Recname2=Trim(request.form("Recname2"))
			If Recname2="" Then
				Response.Write("<script language=""JavaScript"">alert(""���ѣ��������ջ��˵�������"");history.go(-1);</script>")
				response.end
			end if
		Recaddress2=Trim(request.form("Recaddress2"))
			If Recaddress2="" Then
				Response.Write("<script language=""JavaScript"">alert(""���ѣ��������ջ��˵ĵ�ַ��"");history.go(-1);</script>")
				response.end
			end if
		RecPhone2=Trim(request.form("RecPhone2"))
			If RecPhone2="" Then
				Response.Write("<script language=""JavaScript"">alert(""���ѣ��������ջ��˵���ϵ��ʽ��"");history.go(-1);</script>")
				response.end
			end if
		Zipcode2=request.form("ZipCode2")
'==================��ʼ֧����ʽ�����====================
		kuaidi=Trim(request.form("kuaidi"))            '��ݷ�ʽ
		youfei=Cint(Trim(request.form("youfei")))      '�ʷѷ���
		If youfei="" Then
			Response.Write("<script language=""JavaScript"">alert(""���ѣ���������ȷ���ʷѣ�"");history.go(-1);</script>")
			response.end
		end if
		alipay=request.form("alipay")               '֧����ʽ		  
		Transactions=Trim(request.form("Transaction"))  '֧�������ʺ�
		If Transactions="" Then
			Response.Write("<script language=""JavaScript"">alert(""���ѣ�������֧�����ţ�"");history.go(-1);</script>")
			response.end
		end if
		
		set rs1=server.createobject("adodb.recordset")
		sql1="select jiaoyihao from dk501_borderList where jiaoyihao ='"&Transactions&"'"
		rs1.open sql1,conn,1,3
		if not rs1.eof then
				Response.Write("<script language=""JavaScript"">alert(""���ѣ����׺��Ѿ����ڣ�"");history.go(-1);</script>")
				response.end
		else
			Transactions=Trim(request.form("Transaction"))
		end if
'========================================================
jifentotle=request.form("jifensum")
Notes=request.form("Notes")
post=request.form("post")       '���Բ�Ҫ
Gettime=request.form("gettime") '���Բ�Ҫ
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
	inBillNo="dk501-"&yy & mm & dd & xiaoshi & fenzhong & miao & d
	inBillNo1=yy&"��" & mm &"��"& dd &"��"& xiaoshi& "ʱ"& fenzhong &"��"
%>

<%
'===================================��Ʒԭ����̨���õļ۸�====================


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
      <td height="40" background="../image/product/list.jpg" style="padding-left:200px;">
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
      <td><img src="../adv/list.jpg" width="950" height="90" /></td>
    </tr>
    <tr>
      <td height="5"></td>
    </tr>
    <tr>
      <td align="center" class="b-top b-left b-right b-bottom" style="padding:6px 6px 6px 6px;">
	  <table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
        <td height="65" align="left" background="../image/order/news2.jpg" class="b-left b-right">&nbsp;&nbsp;<img src="../image/order/user1.gif" width="11" height="11" /> ��ǰλ�ã�<a href="index.asp">��ҳ</a> ��<a href="shop.asp">��������</a> ���ҵĶ���</td>
      </tr>
      <tr>
        <td height="28" align="left" valign="middle" background="../image/order/bucou1.jpg" class="b-left b-right">&nbsp;&nbsp;&nbsp;<font color="#FFFFFF">�����ύ���1.ѡ���Ʒ������Ʒ 2.��дҪ�ĳ��루��������������Ʒ��3.��д��ϵ��ʽ�Լ����� 4.�ɹ��ύ����</font></td>
  </tr>
      <tr>
        <td height="89" align="left" valign="top" class="b-left b-right b-bottom" style="padding-top:10px; padding-right:10px; padding-bottom:10px; padding-left:10px;"><table width="800" height="50" align="center" cellpadding="0" cellspacing="0" class="jiesuan">
                <tr>
                  <td width="565"><table width="800"  border="0" cellpadding="0" cellspacing="1" bgcolor="#D1D2D3">
                      <tr bgcolor="#FFFFFF">
                        <td width="126" height="25" align="right">���Ķ������Ϊ��</td>
                        <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="6%">&nbsp;</td>
                              <td width="94%" height="25"><span class="bold"><%=inBillNo%></span>����ѯ�������ס�����ţ�</td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td width="126" height="25" align="right">���ύ����ʱ�䣺</td>
                        <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="6%">&nbsp;</td>
                              <td width="94%" height="24"><%=inBillNo1%></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td height="28" align="right" bgcolor="#FFFFFF">�ջ��˵�������&nbsp;</td>
                        <td width="27" bgcolor="#FFFFFF">&nbsp;</td>
                        <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D1D2D3">
                            <tr>
                              <td width="23%" height="25" align="right" bgcolor="#FFFFFF">�ջ��˵�������</td>
                              <td height="25" bgcolor="#FFFFFF"><%=Trim(request.form("Recname2"))%></td>
                            </tr>
                            <tr>
                              <td height="25" align="right" bgcolor="#FFFFFF">�ջ��˵ĵ绰��</td>
                              <td height="25" bgcolor="#FFFFFF"><%=Trim(request.form("RecPhone2"))%></td>
                            </tr>
                            <tr>
                              <td height="25" align="right" bgcolor="#FFFFFF">�ջ��˵ĵ�ַ��</td>
                              <td height="25" bgcolor="#FFFFFF"><%=Trim(request.form("Recaddress2"))%></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td height="28" rowspan="3" align="right">���Ķ�����&nbsp;</td>
                        <td width="27" rowspan="3">&nbsp;</td>
                        <td width="138" height="25" align="right">��Ʒ�ܼƽ�</td>
                        <td width="472" height="25">��<font color="#FF0000"><%=Trim(request.form("sumkiss"))%></font>Ԫ</td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td height="25" align="right">�����ʷѣ�</td>
                        <td width="472" height="25">��<%=Trim(request.form("youfei"))%>Ԫ</td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td height="22" align="right">�ܸ���</td>
                        <td height="22">��<%=formatnum(Trim(request.form("sumkiss"))+Cint(Trim(request.form("youfei"))),2)%></td>
                      </tr>
                  </table></td>
                </tr>
            </table></td>
      </tr>
</table>
	  </td>
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
<!--#include file="../include/login/bootom.asp"-->
</body>
</html>
<% call CloseConn() %>
