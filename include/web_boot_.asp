<hr style=" color:#A10000; width:950px;" />
<table width="950" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#B8B8C0">
  <tr>
    <td height="25" align="center" bgcolor="#FFFFFF" class="Brown">����ָ�� </td>
	<td height="25" align="center" bgcolor="#FFFFFF" class="Brown">�ҵ��ʻ�</td>
    <td height="25" align="center" bgcolor="#FFFFFF" class="Brown">���ʽ</td>
    <td width="160" height="25" align="center" bgcolor="#FFFFFF" class="Brown">��θ���/�˿�</td>
    <td height="25" align="center" bgcolor="#FFFFFF" class="Brown">���ͷ�ʽ</td>
    <td height="25" align="center" bgcolor="#FFFFFF" class="Brown">��������</td>
  </tr>
  <tr>
    <td height="57" align="center" valign="top" bgcolor="#FFFFFF"><a href="Reg.asp">��Աע��</a><br />
      <a href="Contact.asp?id=8">������֪</a><br />
    <a href="Contact.asp?id=10">������ѯ</a></td>
    <td height="57" align="center" valign="top" bgcolor="#FFFFFF"><a href="user_index.asp">�û�����</a><br />
      <a href="Contact.asp?id=14">��Ա����</a><br />
    <a href="Contact.asp?id=13">��Ҫ����</a></td>
    <td height="57" align="center" valign="top" bgcolor="#FFFFFF"><a href="Contact.asp?id=2">����˵��</a><br />
      <a href="Contact.asp?id=12">��������</a></td>
    <td height="57" align="center" valign="top" bgcolor="#FFFFFF"><a href="Contact.asp?id=9">����˿�</a><br />
      <a href="Contact.asp?id=7">��λ���</a></td>
    <td height="57" align="center" valign="top" bgcolor="#FFFFFF"><a href="Contact.asp?id=3">��ͨ���</a><br />
      <a href="Contact.asp?id=4">����ϸ��</a> </td>
    <td height="57" align="center" valign="top" bgcolor="#FFFFFF"><a href="Contact.asp?id=1">��������</a><br />
      <a href="Contact.asp?id=6">��ϵ����</a><a href="/Contact/Contact.asp?id=6"><br />
    </a><a href="Contact.asp?id=12">��������</a></td>
  </tr>
  <tr>
    <td height="20" colspan="6" align="left" valign="top" bgcolor="#FFFFFF" style="padding-top:4px; padding-right:10px; padding-bottom:4px; padding-left:10px;">�������:
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
    <td align="center">�ܹ�˾��<%=adm_tel%> �ƶ��绰��<%=adm_mob%>(����10:00-����19:00)
 <br />
��˾��ַ��<%=adm_address%>�ʱࣺ610000 ICP������<%=adm_icp%><br />
</td>
  </tr>
</table>
