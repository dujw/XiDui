<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<!--#include file="funcation.asp"-->

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
<%
	dim strFileName
	const MaxPerPage=20
	dim totalPut,CurrentPage,TotalPages
	dim i,j
	dim ArticleID
	dim Title
	dim BigClassName,SmallClassName
	dim action
	
%>

<script language=javascript>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="images/menuup.gif";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="images/menudown.gif";
}
}

top.document.title="后台管理"; 
</script>
<%
more=request("more")
action=trim(request("action"))
numbers=trim(request("numbers"))
strFileName="numberslist.asp?more="& more &"&numbers=" & numbers & "&day="&trim(request.QueryString("day"))&""
ID=Request("ID")
%>

<%
	'是否发货物
	if action="edit" then
		id= Trim(Request("ID"))
		set rs=Server.CreateObject("adodb.recordset")
		sql="select * from translist where id="&id
		rs.open sql,conn,1,3
		more=rs("more")
			If more=True then
				sql="update translist set more=False where id="&id
			Else
				sql="update translist set more=True where id="&id
			End if
		set rs=conn.execute(sql)
		Response.Redirect strFileName
	end if
%>


<%

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

'sql="select id,numbers,Company,addr,Consignee,photo,Proxy,dates,more from translist where ID>0"
sql="select * from translist where ID>0 "
if numbers<>"" then
	sql=sql & " and (numbers like '%" & numbers & "%' or photo like '%" & numbers & "%')"	
end if

if trim(request.QueryString("more")) ="1" then
	sql=sql & " and more=1"	
elseif trim(request.QueryString("more")) ="0" then
	sql=sql & " and more=0"
end if

if trim(request.QueryString("day")) <> "" then
	sql=sql & " and dates=#" & trim(request.QueryString("day")) & "# "	
end if

sql=sql & " order by id desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.del.chkAll.checked){
	document.del.chkAll.checked = document.del.chkAll.checked&0;
    } 	
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
  }
function ConfirmDel2()
{
   if(confirm("确定要删除选中的数据吗？一旦删除将不能恢复！"))
     return true;
   else
     return false;
	 
}

</SCRIPT>
<style type="text/css">
<!--
.STYLE1 {color: #FFFFFF}
-->
</style>
<table height="5"><tr><td></td></tr></table>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="862" align="center" valign="top">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#B1BFEE">
          <tr bgcolor="#A4B6D7"> 
            <td height="25" bgcolor="#F4F5FD"><a href="News_Manage.asp">&nbsp;数据资讯管理</a> &gt;&gt; 
            <%
if request.querystring="" then
	response.write "所有数据"
else	
		if numbers<>"" then
			response.write "名称中含有“<font color=blue>" & numbers & "</font>”的数据"
		else
			response.Write("所有数据")
		end if 	
end if
%>
&nbsp;<a href="numberslist.asp?more=1" class="800_hzp_font">已经处理 </a>| <a href="numberslist.asp?more=0" class="800_hzp_font">未处理</a>
</td>
            <td width="150" bgcolor="#F4F5FD">&nbsp; 
            <%
  	if rs.eof and rs.bof then
		response.write "共找到 0 条数据</td></tr></table>"
	else
    	totalPut=rs.recordcount
		if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		end if

    	end if
		response.Write "共找到 " & totalPut & " 条数据"
%></td>
          </tr>
</table><table height="5"><tr><td></td></tr></table>
               <%		
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,false,"条数据"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,false,"条数据"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,false,"条数据"
	    	end if
		end if
	end if
%>
        <%  
sub showContent
   	dim i
    i=0
%>
		<form name="del" method="Post" action="News_Del.asp" onSubmit="return ConfirmDel2();" id="del">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#B1BFEE" class="border" style="word-break:break-all">
          <tr bgcolor="#A4B6D7" class="title" id="imgmenu<%=i%>" onclick="showsubmenu(<%=i%>)"> 
            <td width="26" height="25" align="center" nowrap="nowrap" bgcolor="#F4F5FD">选中</td>
            <td width="48"  height="25" align="center" nowrap="nowrap" bgcolor="#F4F5FD">归属地</td>
            <td width="135" align="center" nowrap="nowrap" bgcolor="#F4F5FD">货物单号</td>
            <td width="82" align="center" bgcolor="#F4F5FD" >买家名字</td>
            <td width="230" align="center" bgcolor="#F4F5FD" >货物描述</td>
            <td width="83" align="center" nowrap="NOWRAP" bgcolor="#F4F5FD" >是否到货</td>
            <td width="107" align="center" nowrap="NOWRAP" bgcolor="#F4F5FD" >是否处理</td>
            <td width="80" align="center" nowrap="NOWRAP" bgcolor="#F4F5FD" >操作</td>
          </tr>
          <%do while not rs.eof%>         
		   <tr class="tdbg"> 
            <td width="26" height="22" align="center" bgcolor="#F4F5FD"> 
             <input name='ID' type='checkbox' onClick="unselectall()" id="ID" value='<%=cstr(rs("ID"))%>'></td>
            <td align="center" nowrap="NOWRAP" bgcolor="#F4F5FD">
			<% 
			if rs("guishu")="1" then 
				response.Write "XIDUI"
			elseif rs("guishu")="2" then
				response.Write "5354"
			else
				response.Write "暂无"
			end if
			%>
			</td>
            <td align="center" nowrap="NOWRAP" bgcolor="#F4F5FD"><%=rs("numbers")%>【<%=rs("Company")%>】</td>
            <td align="center"  bgcolor="#F4F5FD"><%=rs("photo")%></td>
            <td align="center"  bgcolor="#F4F5FD"><%=rs("Consignee")%></td>
            <td align="center"  bgcolor="#F4F5FD"><% if rs("product")=True then %>货已经到
              <%else%><font color="#FF0000">还未到货</font></a>
             <%end if %></td>
            <td width="107" align="center" bgcolor="#F4F5FD">
			<% if rs("more")=True then %>已经处理<%else%><font color="#FF0000">还未处理</font></a> 
		     <%end if %></td>
            <td width="80" align="center" bgcolor="#F4F5FD">
			<a href="#" onclick="showsubmenu(<%=i%>)" id="imgmenu<%=i%>">查看</a> 
			<a href="numbers_modi.asp?ID=<%=rs("id")%>">\处理</a>			</td>
          </tr>
		   <tr id="submenu<%=i%>" style="DISPLAY: none">
		     <td height="22" colspan="8" align="left" bgcolor="#F4F5FD"><span class="user_price">&nbsp;&nbsp;退货描述：</span><%=rs("DU_Note")%> | <span class="user_price">快递方式：</span><%=rs("company")%></td>
	      </tr>
          <%
	i=i+1
	      if i>=MaxPerPage then exit do
	      rs.movenext
	loop
%>
        </table>
	
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="29" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox"></td>
            <td width="221">选中本页显示的所有数据</td>
            <td width="400"><input name="submit" type='submit' value='删除选定的数据'> 
              <input name="Action" type="hidden" id="Action" value="Del"></td>
          </tr>
        </table>
        <%
   end sub   	
%>
      </form>
      <form name="searchsoft" method="post" action="numberslist.asp" onSubmit="return true;" id="searchsoft">
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#B1BFEE" class="border">
        <tr class="tdbg"> 
          
            <td height="30" bgcolor="#F4F5FD"> <strong>&nbsp;查找数据：</strong> <input name="numbers" type="text" class=smallInput id="Title3" size="20" maxlength="50"> 
            <input name="Query" type="submit" id="Query" value="查 询"> &nbsp;&nbsp;请输入数据名称。如果为空，则查找所有数据。</td>
          
        </tr>
    </table></form></td>
  </tr>
</table>
