<%
	
	Dim strSQLServer, strSQLUser, strSQLPassword, strSQLDataBase, strScriptFileSystem
	dim conn
	dim connstr
	dim rs, rSub
	dim sql
'*****************************��װ�����ڴ�����ϵͳ����***********************************************

	strSQLServer="42.121.112.224"			'�������������޸�Ϊ��������SQL�ķ�������
	strSQLUser="bkuser"					'�û��������޸�Ϊ����SQL���û���
	strSQLPassword="bkuser123@#"		'���룬���޸�Ϊ����SQL���û���������û������ӦΪ""
	strSQLDataBase="dk501"				'SQLʹ�õ����ݿ��������޸�Ϊ����SQL�ڵ�ʹ�����ݿ���
	
	
'***************************************************************************************************
	connstr = "driver={SQL Server};server=" & strSQLServer & ";UID=" & strSQLUser & ";PWD=" & strSQLPassword & ";database=" & strSQLDataBase & ""
	set conn = server.createobject("ADODB.CONNECTION")
	if err.number<>0 then
		err.clear
		set conn=nothing
		response.write "<script>window.status=""���� ADO ����ʧ��"";</script>"
		response.end
	else
		conn.open connstr
		if err.number<>0 then
			err.clear
			set conn=nothing
			response.write "<script>window.status=""SQL Server ���ݿ�����ʧ��"";</script>"
					'''�������ݿ��Ƿ���ڣ��Լ� ODBC �Ƿ����
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
