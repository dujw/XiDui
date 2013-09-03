<%
if request.cookies("buyok")("admin")="" then response.write "<meta http-equiv='refresh' content='0;URL=../admin.asp'>"
%>
<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<!--#include file="../../include/md5.asp "-->
<%
sub checkmanage(str)
Set mrs = conn.Execute("select * from manage where username='"&request.cookies("buyok")("admin")&"'")
if not (mrs.bof and mrs.eof) then
manage=mrs("manage")
if instr(manage,str)<=0 then
response.write "<script language='javascript'>"
response.write "alert('警告：您没有此项操作的权限！');"
response.write "location.href='quit.asp';"
response.write "</script>"
response.end
else
session("buyok_admin_login")=0
end if
else 
response.write "<script language='javascript'>"
response.write "alert('没有登陆，不能执行此操作！');"
response.write "location.href='quit.asp';"
response.write "</script>"
response.end
end if
set mrs=nothing
end sub
%>
<%
dim ID,Action,sqlDel,rsDel,FoundErr,ErrMsg,ObjInstalled
ID=trim(request("ID"))
Action=Trim(Request("Action"))
FoundErr=False

if ID="" or Action<>"Del" then
	FoundErr=True	
	Response.write  "<script language=javascript>alert('错误!参数不足!!');history.go(-1);</script>"
	Response.end	
end if
if FoundErr=False then
	if instr(ID,",")>0 then
		dim idarr,i
		idArr=split(ID)
		for i = 0 to ubound(idArr)
			call DelNews(clng(idarr(i)))
		next
	else
		call DelNews(clng(ID))
	end if
end if
if FoundErr=False then
	call CloseConn()
	response.Redirect "numberslist.asp"
else
	call CloseConn()
	call WriteErrMsg()
end if

sub DelNews(ID)
	PurviewChecked=False
	sqlDel="select * from translist where ID=" & CLng(ID)
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3	
		rsDel.delete
		rsDel.update
		set rsDel=nothing
end sub
%>