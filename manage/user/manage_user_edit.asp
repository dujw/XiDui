<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<!--#include file="../../include/md5.asp"-->
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
if action="useredits" then
	call save()
end if 
name=request("id")
set rs=conn.execute("select * from dk501_user where UserId='"&name&"'")
	if rs.eof  then
		response.write "<script language='javascript'>"
		response.write "alert('出错了，没有此用户。');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	usertype=rs("usertype")
	if usertype="2" then
		usertype=type2
		userkou=kou2
	elseif usertype="3" then
		usertype=type3
		userkou=kou3
	elseif usertype="4" then
		usertype=type4
		userkou=kou4
	elseif usertype="5" then
		usertype=type5
		userkou=kou5
	elseif usertype="6" then
		usertype=type6
		userkou=kou6
	else
		usertype=type1
		userkou=kou1
	end if
%>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;查看/编辑会员资料</font></B></font>&nbsp;&nbsp;&nbsp;
      <%
	  	Set rs1 = conn.Execute("select usertype1,usertype2,usertype3,usertype4,usertype4,usertype5,usertype6 from dk501_bconfig")
	    response.Write "<a href=manage_user_list.asp?usertype=1&show=yes>"&rs1("usertype1")&"</a>&nbsp;&nbsp;"
		response.Write "<a href=manage_user_list.asp?usertype=2&show=yes>"&rs1("usertype2")&"</a>&nbsp;&nbsp;"
		response.Write "<a href=manage_user_list.asp?usertype=3&show=yes>"&rs1("usertype3")&"</a>&nbsp;&nbsp;"
		response.Write "<a href=manage_user_list.asp?usertype=4&show=yes>"&rs1("usertype4")&"</a>&nbsp;&nbsp;"
		response.Write "<a href=manage_user_list.asp?usertype=5&show=yes>"&rs1("usertype5")&"</a>&nbsp;&nbsp;"
		response.Write "<a href=manage_user_list.asp?usertype=6&show=yes>"&rs1("usertype6")&"</a>"
	    rs1.close
		set rs1=nothing
	%>
    </TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table width="99%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
  <form name="myinfo" action="manage_user_edit.asp?action=useredits" method="post">
    <tr bgcolor="#FFFFFF">
      <td width='11%' height=22 align=right>会员信息 &nbsp; </td>
      <td height="22" colspan="3"><%=Memo%><%=Name%> | 注册时间：<%=rs("SignDate")%> | 上次登陆时间：<%=rs("lastlogin")%> | 登陆次数：<%=rs("totallogin")%>次 | 上次登陆IP：<%=rs("IP")%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right bgcolor="#EBE9C5">订单数据 &nbsp;</td>
      <td height="22" colspan="3" bgcolor="#EBE9C5">总共卖出：
        <input name="ALL_Total" type="text" id="ALL_Total" value="<%=rs("ALL_Total")%>" size="10" disabled="disabled">
        件| 上月卖出：
        <input name="Month_Total" type="text" id="Month_Total" value="<%=rs("Month_Total")%>" size="10" disabled="disabled">
        单
        <input name="Month_Total_danshu" type="text" id="Month_Total_danshu" value="<%=rs("Month_Total_danshu")%>" size="10" disabled="disabled">
        件 | 上月卖出：
        <input name="Month_danshu" type="text" id="Month_danshu" value="<%=rs("Month_danshu")%>" size="10" disabled="disabled">
        单
        <input name="Current_Total" type="text" id="Current_Total" value="<%=rs("Current_Total")%>" size="10" disabled="disabled">
        件</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right bgcolor="#E9FDDF">消费总金额 &nbsp; </td>
      <td width="33%" height="22" bgcolor="#E9FDDF"><input name="totalsum" type="text" id="totalsum" value="<%=formatnum(rs("totalsum"),2)%>" maxlength=8>
        &nbsp;&nbsp; <img src=../Images/memo.gif alt="订单完成后，会自动追加金额<br>管理员亦可手工修改此处金额" width="16" height="16"> </td>
      <td width="12%" height="22" align="right" bgcolor="#E9FDDF">帐户余额 &nbsp;</td>
      <td width="44%" height="22" bgcolor="#E9FDDF"><input name="yucunkuan" type="text" id="yucunkuan" value="<%=formatnum(rs("yucunkuan"),2)%>" maxlength=8></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right bgcolor="#E9FDDF">会员等级&nbsp;</td>
      <%
	  set rstc=conn.execute("select usertype1,usertype2,usertype3,usertype4,usertype5,usertype6 from dk501_bconfig")
       %>
      <td height="22" bgcolor="#E9FDDF"><select size="1" name="usertype">
          <option value="1" <%if usertype=type1 then%>selected<%end if%>><%=type1%></option>
          <%if type2<>"" then%>
          <option value="2" <%if usertype=type2 then%>selected<%end if%>><%=type2%></option>
          <%if type3<>"" then%>
          <option value="3" <%if usertype=type3 then%>selected<%end if%>><%=type3%></option>
          <%if type4<>"" then%>
          <option value="4" <%if usertype=type4 then%>selected<%end if%>><%=type4%></option>
          <%if type5<>"" then%>
          <option value="5" <%if usertype=type5 then%>selected<%end if%>><%=type5%></option>
          <%if type6<>"" then%>
          <option value="6" <%if usertype=type6 then%>selected<%end if%>><%=type6%></option>
          <%
end if
end if
end if
end if
end if%>
        </select></td>
      <td height="22" align="right" bgcolor="#E9FDDF">积分 &nbsp;</td>
      <td height="22" bgcolor="#E9FDDF"><input name="jifensum" type="text" id="jifensum" value="<%=formatnum(rs("jifensum"),2)%>" maxlength=8>
        &nbsp;&nbsp; <img src=../Images/memo.gif alt="订单完成后，会自动为推荐人追加<br>管理员亦可手工修改此处金额" width="16" height="16"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right>会员密码 &nbsp; </td>
      <td height="22" colspan="3"><input type="text" name="Userpassword" value="" maxlength="16">
        &nbsp;&nbsp; <img src=../Images/memo.gif alt="若不修改密码，请空出" width="16" height="16"> </td>
    </TR>
    <tr bgcolor="#FFFFFF">
      <td width='11%' height=22 align=right bgcolor="#EBE9C5">&nbsp;</td>
      <td height="25" colspan="3" bgcolor="#EBE9C5"><strong>以下是会员的详细信息：</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width='11%' height=22 align=right bgcolor="#F5F5F5">真实姓名 &nbsp; </td>
      <td height="22" bgcolor="#F5F5F5"><input type="text" name="Username" value="<%=rs("UserName")%>" maxlength="16"></td>
      <td height="22" align="right" bgcolor="#F5F5F5">&nbsp;</td>
      <td height="22" bgcolor="#F5F5F5">&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right bgcolor="#F5F5F5">省份 &nbsp;</td>
      <td height="22" bgcolor="#F5F5F5"><input type="text" name="Province" value="<%=rs("Province")%>"></td>
      <td height="22" align="right" bgcolor="#F5F5F5">城市 &nbsp;</td>
      <td height="22" bgcolor="#F5F5F5"><input type="text" name="City" value="<%=rs("City")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right bgcolor="#F5F5F5">联系地址 &nbsp;</td>
      <td height="22" colspan="3" bgcolor="#F5F5F5"><input name="Address" type="text" value="<%=rs("Address")%>" size="80" maxlength="100"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right bgcolor="#F5F5F5">联系电话 &nbsp;</td>
      <td height="22" bgcolor="#F5F5F5"><input type="text" name="HomePhone" value="<%=rs("HomePhone")%>" maxlength="20"></td>
      <td height="22" align="right" bgcolor="#F5F5F5">电子邮箱 &nbsp;</td>
      <td height="22" bgcolor="#F5F5F5"><input type="text" name="UserMail" value="<%=rs("UserMail")%>" maxlength="100"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right bgcolor="#F5F5F5">腾讯ＱＱ &nbsp;</td>
      <td height="22" nowrap bgcolor="#F5F5F5"><input type="text" name="UserQQ" value="<%=rs("UserQQ")%>" maxlength="11">
        &nbsp; &nbsp;
        <%if rs("UserQQ")<>"" then%>
        <a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin=<%=rs("UserQQ")%>&Site=<%=sitename%>&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:<%=rs("UserQQ")%>:6 alt="与该会员交谈或留言"></a>
        <%end if%></td>
      <td height="22" align="right" bgcolor="#F5F5F5">paipai店铺地址 &nbsp;</td>
      <td height="22" bgcolor="#F5F5F5">&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right bgcolor="#F5F5F5">淘宝帐号&nbsp;&nbsp;</td>
      <td height="22" nowrap bgcolor="#F5F5F5"><input name="taobao_name" type="text" id="taobao_name" value="<%=rs("taobao_name")%>" maxlength="11">
        <a target=blank href="http://amos1.taobao.com/msg.ww?v=2&uid=<%=rs("taobao_name")%>&site=cntaobao&s=2" > <img border="0" src="http://amos1.taobao.com/online.ww?v=2&uid=<%=rs("taobao_name")%>&site=cntaobao&s=2" alt="点我咨询售后相关问题" /> </td>
      <td height="22" align="right" bgcolor="#F5F5F5">淘宝店地址&nbsp;&nbsp;</td>
      <td height="22" bgcolor="#F5F5F5"><input name="taobao" type="text" id="taobao" value="<%=rs("taobao")%>" size="30">
        <a href="<%=rs("taobao")%>" target="_blank"><img src=../Images/memo.gif alt="点击查看" width="16" height="16" border="0"></a></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right bgcolor="#F5F5F5">管理附言 &nbsp;</td>
      <td height=22 colspan="3" align=left bgcolor="#F5F5F5"><textarea name="Memo2" cols="80" id="Memo2" style="overflow:auto;"><%=rs("Memo2")%></textarea>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right>是否上货&nbsp; </td>
      <td height="22"><input name="Verification" type="radio" value=1 <%if rs("Verification")<>0 then%>checked<%end if%>>
        <img src="../Images/user_yes.gif" width="22" height="17">上货/通过审核
        <input name="Verification" type="radio" value=0 <%if rs("Verification")=0 then%>checked<%end if%>>
        <img src="../Images/user_no.gif" width="22" height="17">没上货/未审核 </td>
      <td height="22" align="right">审核通过的理由 &nbsp;</td>
      <td height="22"><textarea name="Memo" cols="40" style="overflow:auto;"><%=rs("Memo")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right>会员状态 &nbsp; </td>
      <td height="22"><input name="Status" type="radio" value=1 <%if rs("Status")<>0 then%>checked<%end if%>>
        正常/通过审核
        <input name="Status" type="radio" value=0 <%if rs("Status")=0 then%>checked<%end if%>>
        锁定/未审核</td>
      <td height="22" align="right" bgcolor="#FFFFFF">是否是商品会员 &nbsp;</td>
      <td height="22"><input name="mall" type="radio" value=1 <%if rs("mall")=1 then%>checked<%end if%>>
正常/普通用户
  <input name="mall" type="radio" value=2 <%if rs("mall")=2 then%>checked<%end if%>>
正常/商品用户 </td></tr>
    <tr bgcolor="#FFFFFF">
      <td height=22 align=right>推荐人 &nbsp;</td>
      <td height="22"><font color="#FF0000"><% =rs("tjr") %></font></td>
      <td height="22" align="right" bgcolor="#FFFFFF">时间 &nbsp;</td>
    <td height="22">注册时间（
      <% =rs("SignDate") %>
      ）/最后登录时间（<% =rs("LastLogin") %>）</td></tr>
    <tr bgcolor="#FFFFFF">
      <td colspan=4><input name="id" type="hidden" value="<%=rs("UserId")%>">
        <input name="edit" type="hidden" value="ok">
        <input type="submit" name="Submit" value="确认修改">
        &nbsp;&nbsp; </td>
    </tr>
  </form>
</table>
<%set rs=nothing %>
</BODY>
</html>
<%
sub save()
'保存会员资料
name=request("id")
sql = "select * from dk501_user where UserId='"&name&"'"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,3
if rs.eof and rs.bof then
response.write "<script language='javascript'>"
response.write "alert('出错了，没有此会员。');"
response.write "location.href='manage_user_list.asp?show=yes';"
response.write "</script>"
response.end
rs.close
set rs=nothing
end if
rs("Usertype") = request("usertype")
rs("Username") = trim(request("Username"))
if trim(request("totalsum"))<>"" then rs("totalsum")= trim(request("totalsum"))
rs("UserQQ") = trim(request("UserQQ"))
rs("HomePhone")= trim(request("HomePhone"))
rs("UserMail") = trim(request("UserMail"))
rs("Province") = trim(request("Province"))
rs("City") = trim(request("City"))
rs("Memo") = request("Memo")
rs("Memo2") = request("Memo2")
rs("Status") = Int(request("Status"))
rs("Verification") = Int(request("Verification"))
rs("mall") = Int(request("mall"))
rs("taobao") = trim(request("taobao"))
rs("taobao_name") = trim(request("taobao_name"))
rs("Address") = trim(request("Address"))
if trim(request("Userpassword"))<>"" then rs("Userpassword")=md5(trim(request("Userpassword")))
rs.update
rs.close
set rs = nothing
response.write "<script language='javascript'>"
response.write "alert('操作成功，会员资料已保存！');"
response.write "location.href='manage_user_edit.asp?action=useredit&id="&name&"';"
response.write "</script>"
response.end
end sub


%>
<% call CloseConn() %>
