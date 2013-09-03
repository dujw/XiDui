<%
		call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>")
		end if
		rs.close
		set rs=nothing
Sub PageControl(iCount,pagecount,page,table_style,font_style)
'生成上一页下一页链接
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")

    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
    For Each x In query
        a = Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
            temp = temp & a(0) & "=" & a(1) & "&"
        End If
    Next
    if page<=1 then
        Response.Write ("首页 " & vbCrLf)        
        Response.Write ("上一页 " & vbCrLf)
    else        
        Response.Write("<A HREF=" & action & "?" & temp & "Page=1 class=news>首页</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & " class=news>上一页</A> " & vbCrLf)
    end if
	if page =<5 then				
							for i=1 to 10 
									if page=i then
										response.Write "<font color=#ff0000><del>["&i&"]</del></font>"& vbCrLf
									else
										if i=<pagecount then
											Response.Write("<A HREF=" & action & "?" & temp & "Page="&i&" class=text>["&i&"]</A> " & vbCrLf)
										end if 																
									end if 
							next	
					else
						Response.Write("<A HREF=" & action & "?" & temp & "Page=1 class=text>1...</A> " & vbCrLf)	
						cc=5
						for i=page-4 to page+cc
							if page=i then
								response.Write "<font color=#ff0000>["&i&"]</font>"& vbCrLf
							else  			
								if i=<pagecount then			
									Response.Write("<A HREF=" & action & "?" & temp & "Page="&i&">["&i&"]</A> " & vbCrLf)	
								end if 
							end if 						
						next													
					end if    
    if page>=pagecount then
        Response.Write ("下一页 " & vbCrLf)
        Response.Write ("尾页 " & vbCrLf)            
    else
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & " class=news>下一页</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & pagecount & " class=news>尾页</A> " & vbCrLf)            
    end if
    Response.Write(" 页次：" & page & "/" & pageCount & "页" &  vbCrLf)
    Response.Write(" 共有" & iCount & "条" &  vbCrLf)  
End Sub
%>