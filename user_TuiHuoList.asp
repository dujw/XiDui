<!--#include file="include/connect.asp" -->
<%
call aspsql()
'�ж��û��Ƿ��½
call checklogin()
UserName   = request.cookies("buyok")("userid")
'if request("set")="ok" then call setq()
%>
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
<%
more=request("more")
action=trim(request("action"))
numbers=trim(request("numbers"))
strFileName="user_tuihuolist.asp?more="& more &"&numbers=" & numbers & "&day="&trim(request.QueryString("day"))&""
ID=Request("ID")
%>

<%
	'�Ƿ񷢻���
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
sql="select * from translist where ID>0 and DU_UserName='"&UserName&"' "
if numbers<>"" then
	sql=sql & " and (numbers like '%" & numbers & "%' or time like '%" & numbers & "%')"	
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
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

</script>
</head>
<body>

<!--#include file="include/web_top.asp" -->
<div id="div_Public">
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="260" valign="top"><!--#include file="left.asp"--></td>
    <td valign="top">
	          <table width="690" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="569" height="25" bgcolor="E43A23">&nbsp;<span class="input_search">�û��˻�</span></td>
                  <td width="121" bgcolor="E43A23" align="right"><a href="user_TuiHuo.asp">����˻���Ϣ</a>&nbsp;&nbsp;</td>
                </tr>
                <tr>
                  <td height="5" colspan="2"></td>
                </tr>
                <tr>
                  <td height="5" colspan="2">
                  <!--NN-->
                 
<table height="5"><tr><td></td></tr></table>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="100%" align="center" valign="top">
      <form name="del" method="Post" action="News_Del.asp" onSubmit="return ConfirmDel();">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#B1BFEE">
          <tr bgcolor="#A4B6D7"> 
            <td height="25" bgcolor="#F4F5FD" align="left">&nbsp;�˻��б� &gt;&gt; 
            <%
if request.querystring="" then
	response.write "��������"
else	
		if numbers<>"" then
			response.write "�����к���<font color=blue>" & numbers & "</font>������"
		else
			response.Write("��������")
		end if 	
end if
%>
&nbsp;<a href="user_tuihuolist.asp?more=1" class="800_hzp_font">�Ѿ����� </a>| <a href="user_tuihuolist.asp?more=0" class="800_hzp_font">δ����</a>
</td>
            <td width="150" bgcolor="#F4F5FD">&nbsp; 
            <%
  	if rs.eof and rs.bof then
		response.write "���ҵ� 0 ������</td></tr></table>"
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
		response.Write "���ҵ� " & totalPut & " ������"
%></td>
          </tr>
        </table><table height="5"><tr><td></td></tr></table>
               <%		
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,false,"������"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,false,"������"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,false,"������"
	    	end if
		end if
	end if
%>
        <%  
sub showContent
   	dim i
    i=0
%>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#B1BFEE" class="border" style="word-break:break-all">
          <tr bgcolor="#A4B6D7" class="title" id="imgmenu<%=i%>" onclick="showsubmenu(<%=i%>)"> 
            
            <td width="48"  height="25" align="center" nowrap="nowrap" bgcolor="#F4F5FD">������</td>
            <td width="135" align="center" nowrap="nowrap" bgcolor="#F4F5FD">���ﵥ��</td>
            <td width="82" align="center" bgcolor="#F4F5FD" >�������</td>
            <td width="230" align="center" bgcolor="#F4F5FD" >�˻�����</td>
            <td width="83" align="center" nowrap="NOWRAP" bgcolor="#F4F5FD" >�Ƿ񵽻�</td>
            <td width="107" align="center" nowrap="NOWRAP" bgcolor="#F4F5FD" >�Ƿ���</td>
            <td width="80" align="center" nowrap="NOWRAP" bgcolor="#F4F5FD" >����</td>
          </tr>
          <%do while not rs.eof%>         
		   <tr class="tdbg"> 
            
            <td align="center" nowrap="NOWRAP" bgcolor="#F4F5FD">
			<% 
			if rs("guishu")="1" then 
				response.Write "XIDUI"
			elseif rs("guishu")="2" then
				response.Write "5354"
			else
				response.Write "����"
			end if
			%>
			</td>
            <td align="center" nowrap="NOWRAP" bgcolor="#F4F5FD"><%=rs("numbers")%>��<%=rs("Company")%>��</td>
            <td align="center"  bgcolor="#F4F5FD"><%=rs("photo")%></td>
            <td align="center"  bgcolor="#F4F5FD"><%=rs("DU_Note")%></td>
            <td align="center"  bgcolor="#F4F5FD"><% if rs("product")=True then %>���Ѿ���
              <%else%><font color="#FF0000">��δ����</font></a>
             <%end if %></td>
            <td width="107" align="center" bgcolor="#F4F5FD">
			<% if rs("more")=True then %>�Ѿ�����<%else%><font color="#FF0000">��δ����</font></a> 
		     <%end if %></td>
            <td width="80" align="center" bgcolor="#F4F5FD">
			<a href="#" onclick="showsubmenu(<%=i%>)" id="imgmenu<%=i%>">�鿴���</a> </td>
          </tr>
		   <tr id="submenu<%=i%>" style="DISPLAY: none">
		     <td height="22" colspan="8" align="left" bgcolor="#F4F5FD"><span class="user_price">&nbsp;&nbsp;���������<%=rs("Reason")%> | ��ݷ�ʽ��</span><%=rs("company")%></td>
	      </tr>
          <%
	i=i+1
	      if i>=MaxPerPage then exit do
	      rs.movenext
	loop
%>
        </table>
	
       
        <%
   end sub   	
%>
      </form>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#B1BFEE" class="border">
        <tr class="tdbg"> 
          <form name="searchsoft" method="get" action="user_TuiHuoList.asp">
            <td height="30" bgcolor="#F4F5FD">&nbsp; <a href="user_TuiHuo.asp">����˻���Ϣ</a>&nbsp;&nbsp;</td>
            <td bgcolor="#F4F5FD"><strong>&nbsp;�������ݣ�</strong> <input name="numbers" type="text" class=smallInput id="Title3" size="20" maxlength="50"> 
            <input name="Query" type="submit" id="Query" value="�� ѯ"> &nbsp;&nbsp;�������������ơ����Ϊ�գ�������������ݡ�</td>
          </form>
        </tr>
    </table></td>
  </tr>
</table>

                  <!--NN-->
				  </td>
                </tr>
              </table>
      </td>
  </tr>
</table>

</div>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>