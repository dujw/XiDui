<%
'if lockip<>"0" then
	'if ip<>"" then
	'ip=buyoktxtcheck(ip)
	'lockip=split(cstr(ip),"@")
		'for N=0 to UBound(lockip)
			'if instr(Request.serverVariables("REMOTE_ADDR"),lockip(n))>0 then
				'response.redirect "error.asp?error=99"
				'response.end
			'end if
		'next
	'end if
'end if


'������½X�������������IP
if session("login_error")>=10 then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from dk501_bconfig"
rs.open sql,conn,1,3
userip=Request.serverVariables("REMOTE_ADDR")
if instr(rs("ip"),userip)<0 then rs("ip")=rs("ip")&"@"&userip
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>"
response.write "alert('�����ӷǷ��½��Ա���룬�ѱ�ϵͳ���Ʒ��ʡ�');"
response.write "location.href='login.asp';"
response.write "</script>"
response.end
end if

'ʹ�á�ȡ�����롱ʱ��������X�Σ�����IP
if session("gpw_error")>=10 then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from dk501_bconfig"
rs.open sql,conn,1,3
userip=Request.serverVariables("REMOTE_ADDR")
if instr(rs("ip"),userip)<0 then rs("ip")=rs("ip")&"@"&userip
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>"
response.write "alert('�����ӷǷ��½��Ա���룬���ѱ�ϵͳ���Ʒ��ʡ�');"
response.write "location.href='main.asp';"
response.write "</script>"
response.end
end if

%>
