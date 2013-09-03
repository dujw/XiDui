<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
<script language=javascript>
function CheckAll(form){
for (var i=0;i<form.elements.length;i++){
var e = form.elements[i];
if (e.name != 'chkall') e.checked = form.chkall.checked; 
}
}
</script>
</head>
<BODY>
<%
action=request("action")
if action="del" then 	
	call userdel(request("num"))
end if 
%>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd>
	 &nbsp;
     <a href="manage_user_list.asp?tjr=杜义">推荐人：杜义</a>&nbsp;
     <span style="float:right">
      <%
	  	Set rs1 = conn.Execute("select usertype1,usertype2,usertype3,usertype4,usertype4,usertype5,usertype6 from dk501_bconfig")
	    response.Write "<a href=?usertype=1&show=yes>"&rs1("usertype1")&"</a>&nbsp;&nbsp;"
		response.Write "<a href=?usertype=2&show=yes>"&rs1("usertype2")&"</a>&nbsp;&nbsp;"
		response.Write "<a href=?usertype=3&show=yes>"&rs1("usertype3")&"</a>&nbsp;&nbsp;"
		response.Write "<a href=?usertype=4&show=yes>"&rs1("usertype4")&"</a>&nbsp;&nbsp;"
		response.Write "<a href=?usertype=5&show=yes>"&rs1("usertype5")&"</a>&nbsp;&nbsp;"
		response.Write "<a href=?usertype=6&show=yes>"&rs1("usertype6")&"</a>"
	    rs1.close
		set rs1=nothing
	%>&nbsp;</span>
    </TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table width="99%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
<%
usersql=request.cookies("buyok")("usersql")
if request("show")="yes" then
  response.cookies("buyok")("usersql")=""
end if
pages = 20
'页面数据开始查询开始
keywords=LCase(Trim(request("keywords")))
UserType=Trim(request("UserType"))
GUserType=Trim(request("UserType"))
tjr=Trim(request("tjr"))
'页面数据开始查询结束
set rs = server.createobject("adodb.recordset")
sql="select pid,Verification,Status,UserId,UserName,usertype,ALL_Total,Month_Total,Month_Total_danshu,Month_danshu,Current_Total,taobao,totallogin,SignDate,mall,LastLogin  from dk501_user where pid is not null "
	'Jave.du Do ReportToExcel 2011.04.01 Start
Report_sql="select Province,City,UserName,HomePhone from dk501_user"
	if keywords <> "" then 
		sql = sql & " and UserId like '%" & keywords & "%' "
		Report_sql=Report_sql& " and UserId like '%" & keywords & "%' "
	end if
	if UserType <> "" then 
		sql = sql & " and UserType = '" & UserType & "' "
		Report_sql=Report_sql& " and UserType = '" & UserType & "' "
	end if
	if tjr <> "" then 
		sql = sql & " and tjr = '" & tjr & "' "
		Report_sql=Report_sql& " and tjr = '" & tjr & "' "
	end if
'开始高级搜索用户资料
if request("show")<>"yes" then 
 sql=sql+usersql 
 Report_sql=Report_sql+usersql 
'排序方式
 sql = sql + " order by pid desc"
else
 sql=sql+ " order by pid desc"
end if
Report_sql=Report_sql+ " order by pid desc"

session("excel_sql")=Report_sql
response.write  "<!--"&session("excel_sql")&"-->"
	'Jave.du End
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "<tr bgcolor=#FFFFFF><td colspan=7 align=center height=50>暂时没有会员</td></tr>"
	response.end 
	rs.close
	set rs=nothing
end if
totalrs=rs.RecordCount
rs.pageSize = pages
allPages = rs.pageCount
page = Cint(Request("page"))
If not isNumeric(page) then page=1
if isEmpty(page) or int(page) <= 1 then
	page = 1
elseif int(page) > allPages then
	page = allPages 
end if
rs.AbsolutePage = page
%>
  <tr bgcolor="#f4f5fd">
    <td width=2% height="25" align=center nowrap>选</td>
    <td width=3% align=center nowrap>是否上货</td>
    <td height="25" align=center nowrap>会员ID</td>
    <td height="25" align=center nowrap>会员姓名</td>
    <td height="25" align=center nowrap>会员级别</td>
    <td align=center nowrap>总共卖出</td>
    <td colspan="2" align=center nowrap>上个月卖出</td>
    <td colspan="2" align=center nowrap>本月卖出</td>
    <td height="25" colspan="2" align=center nowrap>店铺地址</td>
    <td height="25" colspan="2" align=center nowrap>注册时间</td>
    <td height="25" align=center nowrap>登陆次数</td>
  </tr>
<%
Total = 0
Do While Not rs.eof and pages>0
	usertype=rs("usertype")
	if usertype="2" then
		usertype=type2
	elseif usertype="3" then
		usertype=type3
	elseif usertype="4" then
		usertype=type4
	elseif usertype="5" then
		usertype=type5
	elseif usertype="6" then
		usertype=type6
	else
		usertype=type1
	end if
%>
<form action=manage_user_list.asp?action=del method=gest enctype="multipart/form-data" name=user>
  <tr bgcolor="#FFFFFF">
    <td height="21" align="center"><input type='checkbox' value='<%=rs("pid")%>' name=num></td>
    <td height="21" align="center"><%
  	if rs("Verification")=1 then 
		response.Write "<img src=../images/user_yes.gif width=22 height=17 alt=此代理已经上货>"
	else
		response.Write "<img src=../images/user_no.gif  width=22 height=17 alt=此代理还未上货>"
	end if
  %></td>
    <td height="21" align="left">
	<a href='manage_user_edit.asp?action=useredit&id=<%=rs("UserId")%>'>
      <% 
if rs("Status")="0" then 
	response.Write "<font color=#FF0000>"&rs("UserId")&"</font>" 
else 
	response.Write rs("UserId") 
end if
%>
    </a><%if rs("mall")=2 then response.Write "<font color=#FF0000>商</font>" end if%>
	</td>
    <td height="21" align="left"><%=rs("UserName")%></td>
    <td height="21" align="center"><%=usertype%></td>
    <td height="21" align="center"><%=rs("ALL_Total")%>件</td>
    <td height="21" align="center" bgcolor="#CAE8F9"><%=rs("Month_Total")%>单</td>
    <td align="center" bgcolor="#F0D0E7"><%=rs("Month_Total_danshu")%>件</td>
    <td height="21" align="center" nowrap><%=rs("Month_danshu")%>单</td>
    <td height="21" align="center" nowrap bgcolor="#00FFFF"><%=rs("Current_Total")%>件</td>
    <td height="21" colspan="2" align="left">&nbsp;<a href="../gotourl.asp?<%=rs("taobao")%>" target=_blank alt="去这个网站逛逛"><%=rs("taobao")%></a></td>
    <td height="21" colspan="2" align="left"><% =rs("SignDate") %></td>
    <td height="21" align=center><%=rs("totallogin")%></td>
  </tr>
  <%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
%>	
  <tr bgcolor="#FFFFFF">
    <td colspan="1" align="center"><input type='checkbox' name=chkall onclick='CheckAll(this.form)'></td>
    <td colspan="14">全选
      <input type=hidden name=action value="del">
      <input name="submit" type=submit onClick="{if(confirm('确认要删除选定的会员吗？')){this.document.user.submit();return true;}return false;}" value="删除"></td>
  </tr>
  </form>
    <tr bgcolor="#FFFFFF">
    <td height="30" colspan="18" align="center">
	<%
rs.close
set rs=nothing
'if allpages<=1 then exit sub
	response.write "总会员数<font color=red>"&totalrs&"</font> &nbsp; 每页<font color=red>20</font>&nbsp; "
if page = 1 then
	response.write "<font color=darkgray>首页 前页</font>"
else
	response.write "<a href=manage_user_list.asp?tjr="&tjr&"&UserType="&GUserType&"&keywords="&keywords&"&Recname2="&Recname2&"&Status="&Status&"&page=1>首页</a> <a href=manage_user_list.asp?tjr="&tjr&"&UserType="&GUserType&"&keywords="&keywords&"&Recname2="&Recname2&"&Status="&Status&"&page="&page-1&">前页</a>"
end if


action="manage_user_list.asp?tjr="&tjr&"&UserType="&GUserType&"&keywords="&keywords&"&Recname2="&Recname2&"&Status="&Status&""				
 				 if page =< 5 then
							for i=1 to 10								
									if page=i then
										response.Write "["&i&"]"& vbCrLf
									else
										if i =< allpages then
											Response.Write("<A HREF=" & action & "&Page="&i&">"&i&"</A> " & vbCrLf)
										end if													
									end if								
							next	
					else
						Response.Write("<A HREF=" & action & "&Page=1>1...</A> " & vbCrLf)	
						cc=5
						for i=page-4 to page+cc
							if page=i then
								response.Write "["&i&"]"& vbCrLf								
							else  			
								if i=< allpages then			
									Response.Write("<A HREF=" & action & "&Page="&i&">"&i&"</A> " & vbCrLf)	
								end if 
							end if 						
						next													
					end if 

if page = allpages then
	response.write "<font color=darkgray> 下页 末页</font>"
else
	response.write " <a href=manage_user_list.asp?tjr="&tjr&"&UserType="&GUserType&"&keywords="&keywords&"&Recname2="&Recname2&"&Status="&Status&"&page="&page+1&">下页</a> <a href=manage_user_list.asp?tjr="&tjr&"&UserType="&GUserType&"&keywords="&keywords&"&Recname2="&Recname2&"&Status="&Status&"&page="&allpages&">末页</a>"
end if
	response.write " &nbsp;&nbsp;&nbsp;第"&page&"页 共"&allpages&"页"
%>	</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="30" colspan="18" align="left"> <!--      <input type="button" name="button" value='1、导出系统内会员的联系方式'  onclick=location.href="ReportUser_to_excel.asp?type=member" disabled>
        <input type="button" name="button" value='2、导出收货人的联系方式'  onclick=location.href="ReportUser_to_excel.asp?type=sell" disabled>
        
        <strong>&nbsp;查找数据：</strong>
          <input name="keywords" type="text" class=smallInput id="keywords" size="30" maxlength="50">
          <input name="Query" type="submit" id="Query" value="查 询">
        &nbsp;&nbsp;-->
  <form name="searchsoft" method="get" action="manage_user_list.asp">        
 <fieldset><legend>&nbsp;附加操作 &nbsp;</legend>
   <table align="center" width="95%" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
      <tr>
	  <td width="20%" align="right" height="30">查询数据：</td>
      <td><input name="keywords" type="text" class=smallInput id="keywords" size="30" maxlength="50">
          <input name="Query" type="submit" id="Query" value="查 询"> &nbsp;请输入用户的ID名称。如果为空，则查找所有数据。</td>
	  </tr>
       <tr>
	  <td width="20%" align="right" height="30">导出数据为Excel文件：</td>
      <td><input type="button" name="button" value='1、根据查询情况导出会员手机号码' <%if request.cookies("buyok")("admin")="杜义" then %>onclick=location.href="ReportUser_to_excel.asp?type=member"<%else%>disabled<%END IF%>>
        <input type="button" name="button" value='2、导出收货人手机号码' <%if request.cookies("buyok")("admin")="杜义" then %>onclick=location.href="ReportUser_to_excel.asp?type=sell"<%else%>disabled<%END IF%>> 
        <input type="button" name="button" value='3、分类快递发货笔数' <%if request.cookies("buyok")("admin")="杜义" then %>onclick=location.href="ReportUser_to_excel.asp?type=express"<%else%>disabled<%END IF%>> 
        </td>
	  </tr>
   </table> 
   </fieldset>
  </form>   
        </td>
  </tr>
</table>
</BODY>
</html>
<%
sub userdel(num)
delnum=num
if delnum="" or isnull(delnum) then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您什么也没有选择！');"
	response.write "location.href='manage_user_list.asp?show=yes';"
	response.write "</script>"
	response.end
end if
	conn.execute("delete from dk501_user where pid in ("&delnum&")")
	response.write "<script language='javascript'>"
	response.write "alert('操作成功，选定的会员已被删除！');"
	response.write "location.href='manage_user_list.asp?show=yes';"
	response.write "</script>"
	response.end
	conn.close
	set conn=nothing
end sub
%>
<% call CloseConn() %>
