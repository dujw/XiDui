<!--#include file="include/connect.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
<% 
SubLar=trim(request("LarCode")) 
SubCode=trim(request("MidCode"))
	function cutstr(tempstr,tempwid)
		if len(tempstr)>tempwid then
			cutstr=left(tempstr,tempwid)&"..."
		else
			cutstr=tempstr
		end if
	end function
%>
</head>
<body>
<!--#include file="include/web_top.asp" -->
<table width="950" height="111" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top"><table width="950" height="96" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="30" bgcolor="#E6EAF3">&nbsp;目前位置 <a href="index.asp">首页</a> 
		   <%
		   if trim(request("name_key")) <> "" then
		   		response.Write " &gt;  您搜索的产品 &gt; " & trim(request("name_key"))
		   else
		   %>
		    &gt; 产品中心
            <%if SubLar <> "" then response.Write " > <a href=product.asp?LarCode="& SubLar &">"& SubLar &"</a>"%>
            <%if SubCode <> "" then response.Write " > <a href=product.asp?LarCode="& SubLar &"&MidCode="& SubCode &">"& SubCode &"</a>"%>
			<% end if%>          </td>
        </tr>
        <tr>
          <td height="18">&nbsp;</td>
        </tr>
        <tr>
          <td><table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td><table width="100%" height="136"  border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="25%" align="left"><%
	sqlprod="select ProdId,ImgPrev,ProdName,Prod1,PriceOrigin,PriceList,PriceList1,PriceList2,PriceList3,PriceList4,PriceList5,PriceList6,AddDate from dk501_product where online=2 "
	if SubLar<>"" then sqlprod = sqlprod + "and Quantity <> 0 and LarCode='"&SubLar&"'"
	if SubCode<>"" then sqlprod = sqlprod + "and MidCode='"&SubCode&"'"
	
	Keywords=trim(request("name_key"))
	if Keywords<>"" then
		sqlprod = sqlprod + " and ProdName like '%"&keywords&"%'"
	end if
	sqlprod = sqlprod + "order by TJDate desc" 
	Set rsprod=Server.CreateObject("ADODB.RecordSet") 
	rsprod.open sqlprod,conn,1,1
	if rsprod.bof and rsprod.eof then
		response.write "<center><br><font color=red>暂无新品,敬请期待！</font>"
		'response.End
	else
	rsprod.PageSize =25 '每页记录条数
	iCount=rsprod.RecordCount '记录总数
	iPageSize=rsprod.PageSize
	maxpage=rsprod.PageCount 
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
	rsprod.AbsolutePage=Page
		if page=maxpage then
			x=iCount-(maxpage-1)*iPageSize
		else
			x=iPageSize
		end if
	For i=1 To x
	%>
	<table width="950" height="136"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <%i=0
	do while not rsprod.eof and (i<rsprod.pagesize)
i=i+1
%>
            <td width="9510" align="left">
                        <table width="25%" align="center" cellpadding="0" cellspacing="0" >
                          <tr>
             <td ><table width="150" height="150" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td align="center" class="b-top b-left b-right b-bottom"><%
if rsprod("ImgPrev")="" then
	response.write "<a href=list.asp?id="&rsprod("ProdId")&" ><img src=pic/none.gif border=0></a>"
else
	response.write "<a href=list.asp?ProdId="&rsprod("ProdId")&"><img border=0 alt='名字："&rsprod("ProdName")&vbCrLf&"市场价："&formatnumber(rsprod("PriceOrigin"),2)&vbCrLf&"零售价："&formatnumber(rsprod("PriceList"),2)&vbCrLf&"上架时间："&split(rsprod("adddate")," ")(0)&"' src='"&rsprod("ImgPrev")&"' width=180 height=180></a>"
end if
%>                                  </td>
                                </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td valign="top" align="right"><table cellspacing="0" cellpadding="0" width="100%" border="0">
                                <tr>
                                  <td colspan="2" align="left" nowrap="nowrap"  class="product_font"><a href="list.asp?ProdId=<%=rsprod("ProdId")%>" target="_blank"><%=cutstr(rsprod("ProdName"),14)%></a> <br />
                                      <font color="#DB3578">尺码：<%=lleft(Trim(rsprod("Prod1")),16)%><br />
                                      </font> 市场价：￥<%=formatnumber(rsprod("PriceOrigin"),2)%><br />
                                      <% if request.cookies("buyok")("userid")<> "" then 
response.Write " <SPAN class=titile>最低售价：￥"&formatnumber(rsprod("PriceList"),2)&"</font></SPAN><br>"
Set rstxt = conn.Execute("select usertype from dk501_user where userid='"&request.cookies("buyok")("userid")&"'") 
if not rstxt.eof then
usertype=Trim(rstxt("usertype"))
end if
set rstxt=nothing
     if usertype="1" then
		response.write ""&type1&"￥"&FormatNum(rsprod("PriceList1"),2)&"<br>"
	 elseif usertype="2" then
		response.write ""&type2&"￥"&FormatNum(rsprod("PriceList2"),2)&"<br>"
	 elseif usertype="3" then
		response.write ""&type3&"￥"&FormatNum(rsprod("PriceList3"),2)&"<br>"
	 elseif usertype="4" then
		response.write ""&type4&"￥"&FormatNum(rsprod("PriceList4"),2)&"<br>"
	 elseif usertype="5" then
		response.write ""&type5&"￥"&FormatNum(rsprod("PriceList5"),2)&"<br>"
	 elseif usertype="6" then
		response.write ""&type6&"￥"&FormatNum(rsprod("PriceList6"),2)&"<br>"
     end if	 
	end if%>
                                    上架时间：<%=split(rsprod("AddDate")," ")(0)%> </td>
                                </tr>
                                <tr>
                                  <td align="left" class="b-bottom"><% if request.cookies("buyok")("userid")="" then %>
                                      <a href="login.asp"><img src="image/product/denglu.gif"  alt="我要登陆" width="57" height="20" border="0" /></a>
                                      <% else%>
                                      <a href="list.asp?ProdId=<%=rsprod("ProdId")%>"><img src="image/product/buy.gif"  alt="我要购买该商品" width="57" height="20" border="0" /></a>
                                      <% end if %>
                                    <a href="user_index.asp" target="_blank"><img src="image/product/fav.gif"  alt="收藏该商品" width="57" height="20"  border="0" /></a> </td>
                                </tr>
                            </table></td>
                          </tr>
                          <tr class=" b-bottom">
                            <td height="10" align="right" valign="top"></td>
                          </tr>
                      </table>					  </td >
            <%
		if (i mod 5=0) and i>=5 then
%>
          </tr>
          <tr >
            <%
	end if
rsprod.movenext
loop
rsprod.close
%>
          </tr>
        </table>					  </td >
                  </tr>
              </table></td >
            </tr>
            <tr>
              <td><%
		next
		call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>")
		end if
		set rsprod=nothing
'注意这里有改动
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

    Response.Write("<table  width=740" & Table_style & ">" & vbCrLf ) 
    Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR>" & vbCrLf )
    Response.Write("<TD align=left>" & vbCrLf )
    Response.Write(font_style & vbCrLf ) 
	
	Response.Write ("&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf)        
    Response.Write(" 共有" & iCount & "种商品" &  vbCrLf)
	Response.Write ("&nbsp;&nbsp;" & vbCrLf)        
    Response.Write(" 页次：" & page & "/" & pageCount & "页" &  vbCrLf)


    if page<=1 then
        Response.Write ("&nbsp;首页 " & vbCrLf)        
        Response.Write ("上一页 " & vbCrLf)
    else       
        Response.Write("<A HREF=" & action & "?" & temp & "Page=1>首页</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">上一页</A> " & vbCrLf)
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
        Response.Write ("末页 " & vbCrLf)            
    else
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">下一页</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">末页</A> " & vbCrLf)            
    end if
	Response.Write("</TD>" & vbCrLf )             
    Response.Write("</TR></form>" & vbCrLf )        
    Response.Write("</table>" & vbCrLf )        
End Sub
%></td >
            </tr>
          </table></td>
        </tr>
    </table></td>
  </tr>
</table>
</td>
</tr>
</table>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>
