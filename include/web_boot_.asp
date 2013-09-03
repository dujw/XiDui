<hr style=" color:#A10000; width:950px;" />
<table width="950" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#B8B8C0">
  <tr>
    <td height="25" align="center" bgcolor="#FFFFFF" class="Brown">新手指南 </td>
	<td height="25" align="center" bgcolor="#FFFFFF" class="Brown">我的帐户</td>
    <td height="25" align="center" bgcolor="#FFFFFF" class="Brown">付款方式</td>
    <td width="160" height="25" align="center" bgcolor="#FFFFFF" class="Brown">如何付款/退款</td>
    <td height="25" align="center" bgcolor="#FFFFFF" class="Brown">配送方式</td>
    <td height="25" align="center" bgcolor="#FFFFFF" class="Brown">帮助中心</td>
  </tr>
  <tr>
    <td height="57" align="center" valign="top" bgcolor="#FFFFFF"><a href="Reg.asp">会员注册</a><br />
      <a href="Contact.asp?id=8">代理须知</a><br />
    <a href="Contact.asp?id=10">批发咨询</a></td>
    <td height="57" align="center" valign="top" bgcolor="#FFFFFF"><a href="user_index.asp">用户中心</a><br />
      <a href="Contact.asp?id=14">会员服务</a><br />
    <a href="Contact.asp?id=13">我要加盟</a></td>
    <td height="57" align="center" valign="top" bgcolor="#FFFFFF"><a href="Contact.asp?id=2">付款说明</a><br />
      <a href="Contact.asp?id=12">帮助中心</a></td>
    <td height="57" align="center" valign="top" bgcolor="#FFFFFF"><a href="Contact.asp?id=9">如何退款</a><br />
      <a href="Contact.asp?id=7">如何换货</a></td>
    <td height="57" align="center" valign="top" bgcolor="#FFFFFF"><a href="Contact.asp?id=3">普通快递</a><br />
      <a href="Contact.asp?id=4">配送细节</a> </td>
    <td height="57" align="center" valign="top" bgcolor="#FFFFFF"><a href="Contact.asp?id=1">关于我们</a><br />
      <a href="Contact.asp?id=6">联系我们</a><a href="/Contact/Contact.asp?id=6"><br />
    </a><a href="Contact.asp?id=12">帮助中心</a></td>
  </tr>
  <tr>
    <td height="20" colspan="6" align="left" valign="top" bgcolor="#FFFFFF" style="padding-top:4px; padding-right:10px; padding-bottom:4px; padding-left:10px;">合作伙伴:
    <%
Set rs_links = conn.Execute("select url,site from dk501_links where online='1' order by txt desc,num asc") 
if not (rs_links.eof and rs_links.bof) then
	Do while not rs_links.eof
	response.write "<a href='"&rs_links("url")&"' target='_blank'>"&lleft(rs_links("site"),12)&"</a>&nbsp;"
	rs_links.movenext
	loop
end if     
rs_links.close
set rs_links=nothing
%></td>
  </tr>
</table>
<div style="display:none">
<script src="http://s13.cnzz.com/stat.php?id=2611623&web_id=2611623" language="JavaScript"></script>
</div>
<table width="950" height="96" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center">总公司：<%=adm_tel%> 移动电话：<%=adm_mob%>(早上10:00-下午19:00)
 <br />
公司地址：<%=adm_address%>邮编：610000 ICP备案：<%=adm_icp%><br />
</td>
  </tr>
</table>
