 <%
 if rs.recordcount=0 then 
 %>
 <br>
 <table width="243" border="0" align="center" cellpadding="0" cellspacing="0" >
		<tr>
		<td height="25" align=center nowrap="nowrap">数据库找不到信息，请查看该文件！</td>
		</tr>
</table>
		 <% else
	  		rs.PageSize =24'每页记录条数
			iCount=rs.RecordCount '记录总数
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
	