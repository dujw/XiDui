<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<%
stype=Trim(Request.QueryString("type"))
'���ݲ�������SQL���
select case stype
 	case "member"
		exeSQL=session("excel_sql")
		Filetitle="��Ա�ֻ�����"
		ToExcel stype,exeSQL,Filetitle
	case "sell"
		exeSQL="select distinct RecName2 as username,RecPhone2 as Mobile from dbo.dk501_borderList"
		Filetitle="�ջ����ֻ�����"
		ToExcel stype,exeSQL,Filetitle
	case "express"
		exeSQL="select year(ordertime)  AS ��,month(ordertime)  AS ��,kuaidi AS ���,count(ordernum) AS ������ from dbo.dk501_borderList  where del<>1 and status='99' group by kuaidi,year(ordertime),month(ordertime) order by ��,��"
		Filetitle="���·�ͳ�Ʒ����ݷ�������"
		ToExcel stype,exeSQL,Filetitle
	case else 
		Response.Write "Error������δ���壡"
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

'���ݲ���----�Զ�����������
	select case stype
 		case "member"
			response.write "ʡ��"&chr(9) &"����"&chr(9) &"����"&chr(9) &"�ֻ�����"&chr(9)&"(�ֻ�����ǰ���'������Ϊ�˷�ֹExcel�г������Զ���ʽ������ʾ�쳣�Ĵ���,���ò����滻����ȥ���ֻ�����ǰ��'����)"&chr(9)&chr(13)
		case "sell"
			response.write "����"&chr(9) &"�ֻ�����"&chr(9)&"(�ֻ�����ǰ���'������Ϊ�˷�ֹExcel�г������Զ���ʽ������ʾ�쳣�Ĵ���,���ò����滻����ȥ���ֻ�����ǰ��'����)"&chr(9)&chr(13)
		case "express"
			response.write "��"&chr(9) &"��"&chr(9)&"��ݹ�˾"&chr(9)&"����������"&chr(9)&chr(13)
		case else 
			Response.Write "Error������δ���壡"
	end select
	
	while not rs.eof
		   i=0
		   while i<cint(total)
		  				 '������ݲ���----���ڳ������ֶε����⴦��
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
'		ErrMsg="�ɹ������ͻ����ϣ�")
'		message 1,ErrMsg,3
		Response.Write session("�ɹ�������")
	wend
		rs.close	
		conn.close 
		Response.Flush
		Response.End 
End Function
%>

 

 
