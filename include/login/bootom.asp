<div id='div_Public'>
<div id='foot_link'><span>������飺</span>
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
%>
</div>
<div id='foot_login'>
<a href="/Contact/Contact.asp?id=1">��������</a> 
<a href="/Contact/Contact.asp?id=2">����˵��</a> 
<a href="/Contact/Contact.asp?id=3">��ͨ���</a> 
<a href="/Contact/Contact.asp?id=4">����ϸ��</a> 
<a href="/Contact/Contact.asp?id=6">��ϵ����</a> 
<a href="/Contact/Contact.asp?id=7">��λ���</a> 
<a href="/Contact/Contact.asp?id=8">������֪</a> 
<a href="/Contact/Contact.asp?id=9">����˿�</a> 
<a href="/Contact/Contact.asp?id=10">������ѯ</a> 
<a href="/Contact/Contact.asp?id=12">��������</a> 
<br>
  Copyright 2008-2010 Tencent.All Rights Reserved , ��Ȩ���� WWW.DK501.COM ����<br>
  �й���Ϣ��ҵ�������ţ���ICP��09003182��
</div>
</div>
