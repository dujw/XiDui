<!--#include file="../include/connect.asp"-->
<%
call checklogin()

	'��ʼ�ж������Ƿ���ڻ���Ƿ����ݣ�
	If Trim(request("ProdId"))="" or Len(Trim(request("ProdId")))=0 then
		response.write "<script language='javascript'>"
		response.write "alert('���棺�����ڷǷ�������');"
		response.write "location.href='shop.asp?prodid="&Trim(request("ProdId"))&"';"
		response.write "</script>"
		response.End()
	end if
	If Trim(request("spec_"&Trim(request("ProdId"))))="" then
		response.write "<script language='javascript'>"
		response.write "alert('��ʾ����ѡ����Ҫ�ĳ��룬лл����֧�֣�');"
		response.write "location.href='shop.asp?prodid="&Trim(request("ProdId"))&"';"
		response.write "</script>"
		response.End()
	end if
	If Trim(request("cart_"&Trim(request("ProdId"))))="" or Trim(request("cart_"&Trim(request("ProdId"))))=0 then
		response.write "<script language='javascript'>"
		response.write "alert('��ʾ������������Ҫ������,�Ҵ�Լ����1��лл����֧�֣�');"
		response.write "location.href='shop.asp?prodid="&Trim(request("ProdId"))&"';"
		response.write "</script>"
		response.End()
	end if

Name=request.cookies("buyok")("userid")
ProdId = Trim(request("ProdId"))
chima  = Trim(request("spec_"&ProdId))
shul   = Cint(Trim(request("cart_"&ProdId)))
jiage  = Cint(Trim(request("Type_"&ProdId)))

'��ʼ�������

		set rsadd=server.createobject("adodb.recordset")
		sql="select * from dk501_border where ProdId='"& ProdId &"' and UserId='"& Name &"' and OrderNum='"& Name &"' and xinhao = '"& chima &"'"
		rsadd.open sql,conn,1,3
		if (rsadd.bof and rsadd.eof) then
			rsadd.AddNew 
					rsadd("OrderNum")=Name
					rsadd("UserId")=Name
					rsadd("ProdId")=ProdId
					rsadd("ProdUnit")=shul
					rsadd("BuyPrice")=jiage
					rsadd("OrderTime")=Now()
					rsadd("jifensum")=0
					rsadd("xzf")=1
					rsadd("xinhao")=chima
					'==================================
				rsadd.Update
			rsadd.close
			set rsadd=nothing
		else
			Response.Write("<script language=""JavaScript"">alert(""��ʾ�����ﳵ���Ѿ��������Ʒ�������ظ��ύ��"");history.go(-1);</script>")
			response.End()
		end if
response.redirect "check.asp"
call CloseConn()
%>