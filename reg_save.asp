<!--#include file="include/connect.asp" -->
<!--#include file="include/md5.asp"-->
<%
call aspsql()
if request("check")<>"reg" then 
	response.redirect "index.asp"
end if
check1=trim(request("check1"))   '输入的验证码
check2=trim(request("check2")) '系统自带的验证码
if check1="" or check1 <> check2 then	
	Response.Write("<script language=""JavaScript"">alert(""错误：请输入正确的验证码，请返回检查！！"");history.go(-1);</script>")
end if
'判断用户是否存在
set rs = conn.execute ("SELECT taobao FROM dk501_user where taobao = '" & Trim(request("taobao")) & "'")
if not (rs.Bof or rs.eof) then				
	response.write "<script language='javascript'>"
	response.write "alert('店铺地址已经有人使用，请输入您其它正确得店铺地址!');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end if
set rs=nothing
			
UserId=Trim(request.form("UserId"))
If Trim(Request.form("UserId"))="" Then
	Response.Write("<script language=""JavaScript"">alert(""请输入用户名！"");history.go(-1);</script>")
	response.end
end if

nowtjr=Trim(request.form("tjr"))

if nowtjr="杜义" then	
		tjr=nowtjr
	else	
		response.write "<script language='javascript'>"
		response.write "alert('错误提示：推荐人错误！请与在线客户联系！');"
		response.write "location.href='javascript:history.go(-1)';"							
		response.write "</script>"	
		response.End()
	end if
	
User_Password=request.form("pw1")
userpassword=md5(user_password)
Username=Trim(request.form("Username"))
HomePhone=Trim(request.form("HomePhone"))
province=Trim(request.form("province"))
city=Trim(request.form("city"))
Address=Trim(request.form("Address"))
taobao=Trim(request.form("taobao"))
taobao_name=request.form("taobao_name ")
taobao_qq=Trim(request.form("taobao_qq"))

jptxt="ゴ|ガ|ギ|グ|ゲ|ザ|ジ|ズ|ヅ|デ|ド|ポ|ベ|プ|ビ|パ|ヴ|ボ|ペ|ブ|ピ|バ|ヂ|ダ|ゾ|ゼ"
jptext=split(jptxt,"|")
for jp=0 to ubound(jptext)
	userid=replace(userid,jptext(jp),"")
	Username=replace(Username,jptext(jp),"")
next
'过滤日文字符及其它常见的非法文字
if request("UserId")<>UserId or request("Username")<>Username then
	response.write "<script language='javascript'>"
	response.write "alert('本系统不允许使用日文字符，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end if

if instr(lcase(userid),"admin")>0 or instr(userid,"管理员")>0  or instr(userid,"游客")>0 then
	'判断用户注册的时候时候是否注册的有含 admin 管理员 游客 等用户名的用户
	response.write "<script language='javascript'>"
	response.write "alert('警告：不允许使用此类用户名，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end if

if buyoktxtcheck(userid)<>userId then
	response.write "<script language='javascript'>"
	response.write "alert('您输入的用户名中含有非法字符，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end if
set rs = conn.execute ("SELECT UserId FROM dk501_user where UserId= '" & UserId & "'")
if not (rs.Bof or rs.eof) then
set rs=nothing
	response.write "<script language='javascript'>"
	response.write "alert('该帐号已经有人使用，请选择其它用户名!');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
else
sqlinfo = "select * from dk501_user"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sqlinfo,conn,1,3
	rs.AddNew
		rs("UserId")=UserId
		rs("UserPassword")=UserPassword
		rs("Username")=Username
		rs("HomePhone")=HomePhone
		rs("Province")=province
		rs("City")=City
		rs("Address")=Address
		rs("IP")=Request.serverVariables("REMOTE_ADDR")
		rs("taobao")=taobao
		rs("taobao_name")=taobao_name
		rs("UserQQ")=taobao_qq
		rs("TJR")=tjr
	rs.Update
rs.close
set rs = nothing
call CloseConn()
response.Redirect("error.asp")

end if
%>
