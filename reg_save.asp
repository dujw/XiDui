<!--#include file="include/connect.asp" -->
<!--#include file="include/md5.asp"-->
<%
call aspsql()
if request("check")<>"reg" then 
	response.redirect "index.asp"
end if
check1=trim(request("check1"))   '�������֤��
check2=trim(request("check2")) 'ϵͳ�Դ�����֤��
if check1="" or check1 <> check2 then	
	Response.Write("<script language=""JavaScript"">alert(""������������ȷ����֤�룬�뷵�ؼ�飡��"");history.go(-1);</script>")
end if
'�ж��û��Ƿ����
set rs = conn.execute ("SELECT taobao FROM dk501_user where taobao = '" & Trim(request("taobao")) & "'")
if not (rs.Bof or rs.eof) then				
	response.write "<script language='javascript'>"
	response.write "alert('���̵�ַ�Ѿ�����ʹ�ã���������������ȷ�õ��̵�ַ!');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end if
set rs=nothing
			
UserId=Trim(request.form("UserId"))
If Trim(Request.form("UserId"))="" Then
	Response.Write("<script language=""JavaScript"">alert(""�������û�����"");history.go(-1);</script>")
	response.end
end if

nowtjr=Trim(request.form("tjr"))

if nowtjr="����" then	
		tjr=nowtjr
	else	
		response.write "<script language='javascript'>"
		response.write "alert('������ʾ���Ƽ��˴����������߿ͻ���ϵ��');"
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

jptxt="��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��"
jptext=split(jptxt,"|")
for jp=0 to ubound(jptext)
	userid=replace(userid,jptext(jp),"")
	Username=replace(Username,jptext(jp),"")
next
'���������ַ������������ķǷ�����
if request("UserId")<>UserId or request("Username")<>Username then
	response.write "<script language='javascript'>"
	response.write "alert('��ϵͳ������ʹ�������ַ���������������룡');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end if

if instr(lcase(userid),"admin")>0 or instr(userid,"����Ա")>0  or instr(userid,"�ο�")>0 then
	'�ж��û�ע���ʱ��ʱ���Ƿ�ע����к� admin ����Ա �ο� ���û������û�
	response.write "<script language='javascript'>"
	response.write "alert('���棺������ʹ�ô����û�����������������룡');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end if

if buyoktxtcheck(userid)<>userId then
	response.write "<script language='javascript'>"
	response.write "alert('��������û����к��зǷ��ַ���������������룡');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end if
set rs = conn.execute ("SELECT UserId FROM dk501_user where UserId= '" & UserId & "'")
if not (rs.Bof or rs.eof) then
set rs=nothing
	response.write "<script language='javascript'>"
	response.write "alert('���ʺ��Ѿ�����ʹ�ã���ѡ�������û���!');"
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
