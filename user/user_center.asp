<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../include/connect.asp"-->
<%
call aspsql()
'�ж��û��Ƿ��½
call checklogin()
Name   = request.cookies("buyok")("userid")
Recname2=Trim(request("Recname2"))
Page=request("page")
	If not isNumeric(page) then
	response.redirect "user_center.asp"
response.end
end if
pagecount=page

if request("cancel")<>"" then call cancelorder()
	'��ʼ�����ؼ�����
	sql = "select OrderNum,OrderTime,RecName2,jiaoyihao,OrderSum,memo,Status,youfei from dk501_borderList where del<>1 and UserId='"&request.cookies("buyok")("userid")&"' "
	if Recname2 <> "" then 
		sql = sql & " and Recname2 like '%" & Recname2 & "%' "
	end if
	sql = sql & " order by OrderTime desc"

set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,1
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="../image/web_css/login.css" rel="stylesheet" type="text/css" />
</head>
<body>
<!--#include file="../include/top.asp"-->
<div id="div_Public">
<table width="950" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="260" valign="top"><!--#include file="left.asp"--></td>
    <td valign="top">
	          <table width="690" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" bgcolor="E43A23">&nbsp;<span class="input_search">�ҵĶ���</span></td>
                </tr>
                <tr>
                  <td height="5"></td>
                </tr>
                <tr>
                  <td height="5"><table width="690"  border="0" align="left" cellpadding="0" cellspacing="0">
                    <tr bgcolor="#E3E2ED" height="30">
                      <td width="120" height="30" align="left"  class="b-bottom">&nbsp;������</td>
                      <td width="120" height="30" align="left"   class="b-bottom">�ύʱ��</td>
                      <td height="30" align="left"   class="b-bottom">�������</td>
                      <td width="80" align="center" valign="middle" class="b-bottom">�ܽ��</td>
                      <td width="180" height="30" align="center" class="b-bottom">��ݵ���</td>
                      <td height="30" align="left"  class="b-bottom">����״̬</td>
                    </tr>
                    <%
if rs.eof and rs.bof  then
	response.write "<tr><td height=30 colspan=8 align=center class=red>��ӭ��½����������ó��վ ��ʱû������Ҫ�Ķ���!</td></tr>"
	else
	  		rs.PageSize =10 'ÿҳ��¼����
			iCount=rs.RecordCount '��¼����
			iPageSize=rs.PageSize
    		maxpage=rs.PageCount 
    		page=cint(request("page"))
    
    if Not IsNumeric(page) or page="" then
        page=1
    else
        page=cint(page)
    end if
    
    if page<1 then
        page=1
    elseif  page>maxpage then
        page=maxpage
    end if
    
    rs.AbsolutePage=Page

	if page=maxpage then
		x=iCount-(maxpage-1)*iPageSize
	else
		x=iPageSize
	end if
	
	i=0
    	do while not rs.eof and (i<rs.pagesize)
	i=i+1

%>
                    <tr bgcolor="#FFFFFF" height="30" onmouseover="this.style.backgroundColor='#FFF7D2'" onmouseout="this.style.backgroundColor=''">
                      <td height="30" align="left" valign="middle" class="b-bottom"><a href="#" onclick="window.open('order_view.asp?OrderNum=<%=rs("OrderNum")%>','blank_','scrollbars=yes,resizable=no,width=630,height=500')"> <%=rs("OrderNum")%></a></td>
                      <td height="30" align="left" valign="middle" class="b-bottom"><%=rs("OrderTime")%></td>
                      <td height="30" align="left" valign="middle" class="b-bottom"><%=rs("RecName2")%></td>
                      <td height="30" align="center" valign="middle" class="b-bottom"><%=formatnum(csng(rs("OrderSum"))+csng(rs("youfei")),2)%></td>
                      <td width="180" height="30" align="left" valign="middle" class="green b-bottom"><% if Trim(rs("memo"))<>"" then response.write rs("memo") else response.Write "�ȴ������У������ĵȴ���" end if %>                      </td>
                      <td height="30" align="left" valign="middle" class="b-bottom"><%
		if rs("Status") = "0" then 
			response.Write "<font color=#FF7E07>�ȴ�������</font>"
		elseif rs("Status") = "11" then
			response.Write "<font color=#7C82AE><s>��Ա����ȡ��</s></font>"
		elseif rs("Status") = "12" then
			response.Write "<font color=#FF00000><s>�޻��˿�ɹ�</s></font>"
		elseif rs("Status") = "2" then 
			response.Write "<font color=#FF00000>�޻������˿�</font>"
		elseif rs("Status") = "3" then
			response.Write "<font color=#FFBB09>�ѷ������տ�<font>"
		elseif rs("Status") = "66" then
			response.Write "<font color=#4A8109>�Ѹ��������</font>"
		elseif rs("Status") = "88" then
			response.Write "<font color=#669900>����֧���ɹ�</font>"
		else
			response.Write "�������"
		end if
		%>                      </td>
                    </tr>
                    <%
	rs.movenext
	loop
	end if
	%>
                  </table></td>
                </tr>
				<tr bgcolor="#FFFFFF" height="30">
        <form name="searchsoft" method="get" action="user_center.asp">
      		<td height="30" class=b-bottom><strong>&nbsp;�������ݣ��������������������֣�</strong>
              <input name="Recname2" type="text" class=uset_input id="Recname2" size="40" maxlength="40">
          <input name="Query" type="submit" id="Query" value="�� ѯ"></td>
    	</form>
        </tr>
              </table>
      <%
		if not (rs.eof and rs.bof) then
		call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>")
		end if
		rs.close
		set rs=nothing
		%></td>
  </tr>
</table>

</div>
<!--#include file="../include/login/bootom.asp"-->
</body>
</html>
<%
Sub PageControl(iCount,pagecount,page,table_style,font_style)
'������һҳ��һҳ����
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")

    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
    For Each x In query
        a = Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
            temp = temp & a(0) & "=" & a(1) & "&"
        End If
    Next

    Response.Write("<table  width=690" & Table_style & ">" & vbCrLf )        
    Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR>" & vbCrLf )
    Response.Write("<TD align=left width=900>" & vbCrLf )
    Response.Write(font_style & vbCrLf ) 
    Response.Write("����" & iCount & "������" &  vbCrLf)
    Response.Write(" ҳ�Σ�" & page & "/" & pageCount & "ҳ " &  vbCrLf)


    if page<=1 then
        Response.Write ("��ҳ " & vbCrLf)        
        Response.Write ("��ҳ " & vbCrLf)
    else       
        Response.Write("<A HREF=" & action & "?" & temp & "Page=1>��ҳ</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">��ҳ</A> " & vbCrLf)
    end if

			if Page =< 5 then
							for i=1 to 10								
									if Page=i then
										response.Write "["&i&"]"& vbCrLf
									else
										if i =< maxpage then
											Response.Write("<A HREF=" & action & "?" & temp & "Page="&i&">"&i&"</A> " & vbCrLf)
										end if													
									end if								
							next	
					else
						Response.Write("<A HREF=" & action & "?" & temp & "Page=1>1...</A> " & vbCrLf)	
						cc=5
						for i=page-4 to page+cc
							if page=i then
								response.Write "["&i&"]"& vbCrLf								
							else  			
								if i=< maxpage then			
									Response.Write("<A HREF=" & action & "?" & temp & "Page="&i&">"&i&"</A> " & vbCrLf)	
								end if 
							end if 						
						next													
					end if
					
    if page>=pagecount then
        Response.Write ("��ҳ " & vbCrLf)
        Response.Write ("ĩҳ " & vbCrLf)            
    else
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">��ҳ</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">ĩҳ</A> " & vbCrLf)            
    end if   
    Response.Write("</TD>" & vbCrLf )                
    Response.Write("</TR>" & vbCrLf )        
    Response.Write("</table>" & vbCrLf )        
End Sub



%>
<% call CloseConn() %>
