<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<%
stype=Trim(Request.QueryString("type"))
'根据参数创建SQL语句
select case stype
 	case "member"
		exeSQL=session("excel_sql")
		Filetitle="会员手机号码"
		ToExcel stype,exeSQL,Filetitle
	case "sell"
		exeSQL="select distinct RecName2 as username,RecPhone2 as Mobile from dbo.dk501_borderList"
		Filetitle="收货人手机号码"
		ToExcel stype,exeSQL,Filetitle
	case "express"
		exeSQL="select year(ordertime)  AS 年,month(ordertime)  AS 月,kuaidi AS 快递,count(ordernum) AS 订单数 from dbo.dk501_borderList  where del<>1 and status='99' group by kuaidi,year(ordertime),month(ordertime) order by 年,月"
		Filetitle="按月份统计分类快递发货笔数"
		ToExcel stype,exeSQL,Filetitle
	case else 
		Response.Write "Error：参数未定义！"
end select
'  
Function ToExcel(stype,exeSQL,Filetitles)
	Response.Clear
	Response.ContentType = "text/xls"
	Response.AddHeader "content-disposition", "attachment; filename="&Filetitles&replace(replace(replace(trim(now()),"-",""),":","")," ","")&".xls"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	Set rs=conn.execute(exeSQL)
	'Set rs=conn.execute(session("excel_sql"))
	total=rs.fields.count

'根据参数----自定义首行列名
	select case stype
 		case "member"
			response.write "省份"&chr(9) &"城市"&chr(9) &"姓名"&chr(9) &"手机号码"&chr(9)&"(手机号码前面的'符号是为了防止Excel中长数字自动格式化后显示异常的处理,请用查找替换功能去掉手机号码前的'符号)"&chr(9)&chr(13)
		case "sell"
			response.write "姓名"&chr(9) &"手机号码"&chr(9)&"(手机号码前面的'符号是为了防止Excel中长数字自动格式化后显示异常的处理,请用查找替换功能去掉手机号码前的'符号)"&chr(9)&chr(13)
		case "express"
			response.write "年"&chr(9) &"月"&chr(9)&"快递公司"&chr(9)&"发货单数量"&chr(9)&chr(13)
		case else 
			Response.Write "Error：参数未定义！"
	end select
	
	while not rs.eof
		   i=0
		   while i<cint(total)
		  				 '按需根据参数----列内长数字字段的特殊处理
		   		         select case stype
							  case "member"
								  if i=3 Then
									   Data=Data&"'"&rs(i)&chr(9)
								   else
									  Data=Data&rs(i)&chr(9)
								  End if	
							  case "sell"
								  if i=1 Then
									   Data=Data&"'"&rs(i)&chr(9)
								   else
									  Data=Data&rs(i)&chr(9)
								  End if
							case "express"
									  Data=Data&rs(i)&chr(9)
						  end select						 
					i=i+1  	
		   wend
		Response.Write Data&chr(13)
		Data=""
		rs.moveNext
'		ErrMsg="成功导出客户资料！")
'		message 1,ErrMsg,3
		Response.Write session("成功导出！")
	wend
		rs.close	
		conn.close 
		Response.Flush
		Response.End 
End Function
%>

 

 
