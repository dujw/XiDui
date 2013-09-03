<!--#include file="../include/connect.asp"-->
<%
call checklogin()

	'开始判断数据是否存在或则非法数据！
	If Trim(request("ProdId"))="" or Len(Trim(request("ProdId")))=0 then
		response.write "<script language='javascript'>"
		response.write "alert('警告：您属于非法操作！');"
		response.write "location.href='shop.asp?prodid="&Trim(request("ProdId"))&"';"
		response.write "</script>"
		response.End()
	end if
	If Trim(request("spec_"&Trim(request("ProdId"))))="" then
		response.write "<script language='javascript'>"
		response.write "alert('提示：请选择你要的尺码，谢谢您的支持！');"
		response.write "location.href='shop.asp?prodid="&Trim(request("ProdId"))&"';"
		response.write "</script>"
		response.End()
	end if
	If Trim(request("cart_"&Trim(request("ProdId"))))="" or Trim(request("cart_"&Trim(request("ProdId"))))=0 then
		response.write "<script language='javascript'>"
		response.write "alert('提示：请输入你想要的数量,且大约等于1，谢谢您的支持！');"
		response.write "location.href='shop.asp?prodid="&Trim(request("ProdId"))&"';"
		response.write "</script>"
		response.End()
	end if

Name=request.cookies("buyok")("userid")
ProdId = Trim(request("ProdId"))
chima  = Trim(request("spec_"&ProdId))
shul   = Cint(Trim(request("cart_"&ProdId)))
jiage  = Cint(Trim(request("Type_"&ProdId)))

'开始添加数据

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
			Response.Write("<script language=""JavaScript"">alert(""提示：购物车中已经有这款商品，请勿重复提交！"");history.go(-1);</script>")
			response.End()
		end if
response.redirect "check.asp"
call CloseConn()
%>