 <%
 if rs.recordcount=0 then 
 %>
 <br>
 <table width="243" border="0" align="center" cellpadding="0" cellspacing="0" >
		<tr>
		<td height="25" align=center nowrap="nowrap">���ݿ��Ҳ�����Ϣ����鿴���ļ���</td>
		</tr>
</table>
		 <% else
	  		rs.PageSize =24'ÿҳ��¼����
			iCount=rs.RecordCount '��¼����
			iPageSize=rs.PageSize
    		maxpage=rs.PageCount 
    		page=request("page")
    
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
	%>
	