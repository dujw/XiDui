<%
	
	Dim strSQLServer, strSQLUser, strSQLPassword, strSQLDataBase, strScriptFileSystem
	dim conn
	dim connstr
	dim rs, rSub
	dim sql
'*****************************安装后请在此设置系统变量***********************************************

	strSQLServer="42.121.112.224"			'服务器名，请修改为你主机中SQL的服务器名
	strSQLUser="bkuser"					'用户名，请修改为你在SQL的用户名
	strSQLPassword="bkuser123@#"		'密码，请修改为你在SQL的用户的密码如没设密码应为""
	strSQLDataBase="dk501"				'SQL使用的数据库名，请修改为你在SQL在的使用数据库名
	
	
'***************************************************************************************************
	connstr = "driver={SQL Server};server=" & strSQLServer & ";UID=" & strSQLUser & ";PWD=" & strSQLPassword & ";database=" & strSQLDataBase & ""
	set conn = server.createobject("ADODB.CONNECTION")
	if err.number<>0 then
		err.clear
		set conn=nothing
		response.write "<script>window.status=""创建 ADO 对象失败"";</script>"
		response.end
	else
		conn.open connstr
		if err.number<>0 then
			err.clear
			set conn=nothing
			response.write "<script>window.status=""SQL Server 数据库连接失败"";</script>"
					'''请检查数据库是否存在，以及 ODBC 是否存在
			response.end
		end if
	end if
	
%>
<!--#include file="config.asp" -->
<!--#include file="../adv/config.asp" -->
<!--#include file="buyok.asp" -->
<!--#include file="functions.asp" -->
<%
sub CloseConn()
	conn.close
	set conn=nothing
end sub
%>
