<%
' ============================================
' 格式化时间(显示)
' 参数：n_Flag
' 1:"yyyy-mm-dd hh:mm:ss"
' 2:"yyyy-mm-dd"
' 3:"hh:mm:ss"
' 4:"yyyy年mm月dd日"
' 5:"yyyymmdd"
'   6:"yyyymmddhhmmss" 
' ============================================
Function Format_Time(s_Time, n_Flag)
 Dim y, m, d, h, mi, s
 Format_Time = ""
 If IsDate(s_Time) = False Then Exit Function
 y = cstr(year(s_Time))
 m = cstr(month(s_Time))
 If len(m) = 1 Then m = "0" & m
 d = cstr(day(s_Time))
 If len(d) = 1 Then d = "0" & d
 h = cstr(hour(s_Time))
 If len(h) = 1 Then h = "0" & h
 mi = cstr(minute(s_Time))
 If len(mi) = 1 Then mi = "0" & mi
 s = cstr(second(s_Time))
 If len(s) = 1 Then s = "0" & s
 Select Case n_Flag
 Case 1
  ' yyyy-mm-dd hh:mm:ss
  Format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
 Case 2
  ' yyyy-mm-dd
  Format_Time = y & "-" & m & "-" & d
 Case 3
  ' hh:mm:ss
  Format_Time = h & ":" & mi & ":" & s
 Case 4
  ' yyyy年mm月dd日
  Format_Time = y & "年" & m & "月" & d & "日"
 Case 5
  ' yyyymmdd
  Format_Time = y & m & d
 case 6
  'yyyymmddhhmmss
  format_time= y & m & d & h & mi & s
 End Select
End Function

function lleft(content,lef)
	for le=1 to len(content)
		if asc(mid(content,le,1))<0 then
			lef=lef-2
		else
			lef=lef-1
		end if
		if lef<=0 then exit for
	next
	lleft=left(content,le)
end function
function llen(content)
	truelen=0
	for le=1 to len(content)
	if asc(mid(content,le,1))<0 then
		truelen=truelen+2
	else
		truelen=truelen+1
	end if
next
llen=truelen
end function
	function FormatNum(num,n)
	if num<1 then
		num="0"&cstr(FormatNumber(num,n))
	else
		num=cstr(FormatNumber(num,n))
	end if
	FormatNum=replace(num,",","")
end function
function buyoktxtcheck(txt)
buyoktxtcheck=txt
chrtxt="33|34|35|36|37|38|39|40|41|42|43|44|47|58|59|60|61|62|63|91|92|93|94|96|123|124|125|126|128"
chrtext=split(chrtxt,"|")
for c=0 to ubound(chrtext)
buyoktxtcheck=replace(buyoktxtcheck,chr(chrtext(c)),"")
next
end function
function ii11ii1(ii1liil)
ii11iil=split(ii1liil,".")
ii11ii1=ii11iil(0)&"."&ii11iil(1)&"."&ii11iil(2)&".**"
end function

sub checklogin()
set rscheck=conn.execute("select UserId from dk501_user where UserId='"&request.cookies("buyok")("userid")&"'")
if rscheck.eof and rscheck.bof then
	response.write "<meta http-equiv='refresh' content='0;URL=login.asp'>"
	response.end
end if
set rscheck=nothing
end sub

function checkuserkou()
	if request.cookies("buyok")("userid")="" then
		checkuserkou=10
	else
		checkuserkou=request.cookies("buyok")("userkou")
	if request.cookies("buyok")("userkou")="" then checkuserkou=10
	end if
end function
sub checkmanage(str)
Set mrs = conn.Execute("select * from manage_user where username='"&request.cookies("buyok")("admin")&"'")
if not (mrs.bof and mrs.eof) then
manage=mrs("manage")
	if instr(manage,str)<=0 then
		response.write "<script language='javascript'>"
		response.write "alert('警告：您没有此项操作的权限！');"
		response.write "location.href='quit.asp;"
		response.write "</script>"
		response.end
	else
		session("buyok_admin_login")=0
	end if
else 
	response.write "<script language='javascript'>"
	response.write "alert('没有登陆，不能执行此操作！');"
	response.write "location.href='/manage/quit.asp';"
	response.write "</script>"
	response.end
end if
set mrs=nothing
end sub
sub aspsql()
SQL_injdata = "'|;|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare"
SQL_inj = split(SQL_Injdata,"|")
	If Request.Form<>"" Then
		For Each Sql_Post In Request.Form
			For SQL_Data=0 To Ubound(SQL_inj)
				if instr(Request.Form(Sql_Post),Sql_Inj(Sql_DATA))>0 Then
					response.write "<script language='javascript'>"
					response.write "alert('网站安全提示：请不要在参数中包含非法字符！');"
					response.write "location.href='javascript:history.go(-1)';"
					response.write "</script>"
					response.end
				end if
			next
		next
	end if
	If Request.QueryString<>"" Then
		For Each SQL_Get In Request.QueryString
			For SQL_Data=0 To Ubound(SQL_inj)
				if instr(Request.QueryString(SQL_Get),Sql_Inj(Sql_DATA))>0 Then
					response.write "<script language='javascript'>"
					response.write "alert('网站安全提示：请不要在参数中包含非法字符！');"
					response.write "location.href='javascript:history.go(-1)';"
					response.write "</script>"
					response.end
				end if
			next
		Next
	end If
end sub
%>