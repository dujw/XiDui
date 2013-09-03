<%
if request.cookies("buyok")("admin")="" then response.write "<meta http-equiv='refresh' content='0;URL=../admin.asp'>"
%>

<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<%
sub checkmanage(str)
Set mrs = conn.Execute("select * from manage where username='"&request.cookies("buyok")("admin")&"'")
if not (mrs.bof and mrs.eof) then
manage=mrs("manage")
if instr(manage,str)<=0 then
response.write "<script language='javascript'>"
response.write "alert('警告：您没有此项操作的权限！');"
response.write "location.href='quit.asp';"
response.write "</script>"
response.end
else
session("buyok_admin_login")=0
end if
else 
response.write "<script language='javascript'>"
response.write "alert('没有登陆，不能执行此操作！');"
response.write "location.href='quit.asp';"
response.write "</script>"
response.end
end if
set mrs=nothing
end sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
<HEAD>
  <title>货物单号系统</title>  
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
  <script language="javascript" src="js/area.js"></script>
</HEAD>
<BODY bgcolor="#D9EAFC">
<%
dim count
dim action
action=trim(request("action"))
%>

<%
if request("no")="modi" then
newsid=request("newsid")
	numbers=Trim(request("numbers"))  '货物单号	
	If Trim(Request.form("numbers"))="" Then
		Response.Write("<script language=""JavaScript"">alert(""错误：请输入货物单号！！"");history.go(-1);</script>")
		response.end
	end if
	Company=Trim(request("Company"))        '快递方式
	Consignee=Trim(request("Consignee"))	'收货物人的姓名
	photo=Trim(request("photo"))'收货物人的电话	
	product=request.form("product") '是否收到货物
	more=request.form("more")       '是否已经处理
	guishu=request.form("guishu") 
		if more="0" then
			Reason=""
		else
			Reason=Trim(request.form("Reason"))
				If Trim(Request.form("Reason"))="" Then
				Response.Write("<script language=""JavaScript"">alert(""错误：请输入没发货的详细介绍！"");history.go(-1);</script>")
				response.end
			end if
		end if 	

set rs=server.createobject("adodb.recordset")
sql="select * from translist where id="&newsid
rs.open sql,conn,1,3
			rs("numbers")=numbers
			rs("Company")=Company
			rs("Consignee")=Consignee
			rs("photo")=photo
			rs("product")=product
			rs("more")=more
			rs("Reason")=Reason	
			rs("guishu")=guishu	
				if more = 1 then
					rs("time")=now()&" | "&request.cookies("buyok")("admin")
				end if	
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('数据修改成功！');" & Chr(13)
		response.write "window.document.location.href='numberslist.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
end if%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=750 border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  
    <TR> 
      <TD bgColor=#f4f5fd height=30>　<B><font color="#FF0000">&nbsp;货运单号码修改</font></B></TD>
    </TR>
  
</TABLE>
<% 
		newsid=request("id")
		Set rso=Server.CreateObject("ADODB.RecordSet") 
		sql="select * from translist where  id="&newsid
		rso.Open sql,conn,1,1
			if rso.eof and rso.bof then
				response.Write("没有记录")				
			else
			Company=Trim(rso("Company"))
	%> 
<table height="5"><tr><td></td></tr></table>
   <table width="750"  border="0" cellpadding="5" cellspacing="1" bgcolor="#b1bfee">
     <form name="addNEWS" method="post" action="?no=modi">
       <TR bgcolor="#FFFFFF">
         <TD width="12%" nowrap>货运单号：             </TD>
         <TD colspan="2"><input name="numbers" type="text" id="numbers" value="<%=rso("numbers")%>" maxlength="20">
         例如：2007282627 (<span class="800_hzp_red">没有的请选择其它快递</span>)发货单位   
		       
           <select name="Company" id="Company">
             <option value="申通" <% if Company = "申通" then response.Write "selected" end if %>>== 申通 ==</option>
             <option value="圆通" <% if Company = "圆通" then response.Write "selected" end if %>>== 圆通 ==</option>
             <option value="顺丰" <% if Company = "顺丰" then response.Write "selected" end if %>>== 顺丰 ==</option>
             <option value="EMS"  <% if Company = "EMS" then response.Write "selected" end if %>>== EMS ==</option>
			 <option value="中通" <% if Company = "中通" then response.Write "selected" end if %>>== 中通 ==</option>
			 <option value="天天" <% if Company = "天天" then response.Write "selected" end if %>>== 天天 ==</option>
			 <option value="中诚" <% if Company = "中诚" then response.Write "selected" end if %>>== 中诚 ==</option>
			 <option value="韵达" <% if Company = "韵达" then response.Write "selected" end if %>>== 韵达 ==</option>
			 <option value="其它" <% if Company = "其它" then response.Write "selected" end if %>>== 其它 ==</option>
           </select>         </TD>
       </TR> 
       <TR bgcolor="#FFFFFF">
         <TD nowrap>邮货人名字：           </TD>
         <TD><input name="photo" type="text" id="photo" value="<%=rso("photo")%>" size="50"></TD>
          <TD width="48%">代理ID：<b><font color="#FF0000"><a href="../user/manage_user_edit.asp?action=useredit&id=<%=rso("DU_UserName")%>"><%=rso("DU_UserName")%></a></font></b>
          </TD>
       </TR>
	   <TR bgcolor="#FFFFFF">
         <TD nowrap>货物描述：         </TD>
         <TD colspan="2"><textarea name="Consignee" cols="50" rows="4" id="Consignee"><%=rso("Consignee")%></textarea></TD>
       </TR>
       <TR bgcolor="#FFFFFF">
         <TD nowrap>退货详情：         </TD>
         <TD colspan="2"><textarea name="Consignee" cols="50" rows="4" id="Consignee" disabled><%=rso("DU_Note")%></textarea></TD>
       </TR>
       <TR bgcolor="#FFFFFF">
         <TD nowrap>货物归属：</TD>
         <TD colspan="2"><input name="guishu"  type="radio" value=1 <%if rso("guishu")=1 then Response.Write "checked"%>  />蛋壳网</TD>
       </TR>
	   <TR bgcolor="#FFFFFF">
         <TD nowrap>是否收到货物：</TD>
         <TD colspan="2"><input name="product"  type="radio" value=False <%if rso("product")=False then Response.Write "checked"%>  />
否
  <input name="product"  type="radio"  value=True <%if rso("product")=True then Response.Write "checked"%>/>
是</TD>
       </TR>
	   <TR bgcolor="#FFFFFF">
         <TD nowrap>是否已经处理：
           <script language=javascript>function show(obj)
			{
				if(document.getElementById("r1").checked)
					{
					document.getElementById("PhotoUrl").style.display=''
					}
				else
					{
					document.getElementById("PhotoUrl").style.display='none'
					}
			}
			</script></TD>
         <TD colspan="2"><input name="more"  type="radio" onClick="show()" value="0" <%if rso("more")=False  then Response.Write "checked"%>/>
否
  <input name="more"  type="radio" onClick="show()" value="1" id="r1" <%if rso("more")=True  then Response.Write "checked"%>/><%=rso("more")%>
是
<script language=javascript>
			</script></TD>
       </TR>
	   <tr bgcolor="#ECF5FF"  id="PhotoUrl"  <%if rso("more")="False" then %>style="display:none;"<% end if %>>
         <TD nowrap>处理意见：</TD>
         <TD width="40%"><textarea name="Reason" cols="50" rows="5" id="Reason"><%=rso("Reason")%></textarea></TD>
	     <TD width="48%"><font color="#FF0000">============模板====<br>
	     </font><font color="#FF0000">货号:  尺码:  有无损坏:<br>
请代理联系在先QQ或旺旺<br>
	     ====================</font></TD>
	   </TR>
	   <tr bgcolor="#ECF5FF">
         <TD nowrap>收到货的时间</TD>
         <TD colspan="2"><%=rso("addtime")%></TD>
       </tr>
	   <tr bgcolor="#ECF5FF">
         <TD nowrap>处理货的时间</TD>
         <TD colspan="2"><%=rso("time")%></TD>
       </tr>	  

       <TR bgcolor="#FFFFFF">
         <TD colspan="3">		
		 <input type="submit" name="submit" value="修改完成">
		 <input type="hidden" name="newsId" value="<%=newsId%>">
		 <input type="reset" name="Submit2" value="重新修改"></TD>
       </TR>
     </form>
</TABLE> 
         <% End If
rso.close
set rso=nothing
 %>
</body>
</HTML>
